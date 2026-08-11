"""
محرك تتبع الدفعات (Batch/FIFO Engine) — منقول بنفس منطق كود Google Apps Script بتاع المستخدم
الفكرة: كل حركة استلام = "دفعة" (batch) منفصلة، وكل صرف بيتخصم من أقدم دفعة متاحة (FIFO)
التحويلات بين المواقع (Transfer order shipment/receive) بتتربط ببعض عن طريق رقم المستند (Number)
"""
import pandas as pd
import numpy as np


def fix_transfer_dates(df: pd.DataFrame) -> pd.DataFrame:
    """
    تصحيح تاريخ 'Transfer order receive' ليطابق تاريخ الـ 'Transfer order shipment' المقابل له
    (نفس منطق fixExactRowTransferDates بتاعتك) — عشان الدفعة تاخد تاريخها الحقيقي
    مش تاريخ وصول الورقة للمخزن التاني، وده اللي بيمنع ظهور "رصيد سالب مؤقت" غلط
    """
    df = df.copy()
    df["_key"] = df["Item number"].astype(str) + "|" + df["Variant number"].astype(str) + "|" + df["Number"].astype(str)
    df["_ref_lower"] = df["Reference"].astype(str).str.lower()

    shipment_map = {}
    for key, sub in df[df["_ref_lower"].str.contains("shipment", na=False)].groupby("_key"):
        shipment_map[key] = list(sub.index)

    for idx, row in df[df["_ref_lower"].str.contains("receive", na=False)].iterrows():
        key = row["_key"]
        if shipment_map.get(key):
            ship_idx = shipment_map[key].pop(0)
            df.at[idx, "Physical date"] = df.at[ship_idx, "Physical date"]

    return df.drop(columns=["_key", "_ref_lower"])


def get_status(age_days: float, consumption_ratio: float) -> str:
    if age_days <= 30 and consumption_ratio == 0:
        return "🟢 Fresh Batch"
    if age_days >= 180 and consumption_ratio < 0.05:
        return "💣 Dead Batch"
    if age_days >= 90 and consumption_ratio == 0:
        return "💣 Idle Stock"
    if age_days >= 90 and consumption_ratio < 0.05:
        return "⚠ Slow Moving"
    if consumption_ratio < 0.05:
        return "🟡 Low Consumption"
    return "🟢 Active Batch"


def run_fifo_engine(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
    """
    القلب الرياضي: بيعالج الحركات بالترتيب الزمني، ويبني دفعات (batches)،
    يخصم الصرف بنظام FIFO، ويربط التحويلات ببعضها بنفس منطق الكود الأصلي
    """
    df = fix_transfer_dates(df)
    df["Physical date"] = pd.to_datetime(df["Physical date"], errors="coerce")
    df = df.dropna(subset=["Physical date"])

    df["_key"] = df["Item number"].astype(str) + "|" + df["Variant number"].astype(str)
    df["_loc"] = np.where(
        df["Warehouse"].notna() & (df["Warehouse"].astype(str) != ""),
        df["Site"].astype(str) + "-" + df["Warehouse"].astype(str),
        df["Site"].astype(str),
    )
    df["_ref"] = df["Reference"].astype(str).str.strip()
    df["_ref_lower"] = df["_ref"].str.lower()

    # ترتيب: بالتاريخ، ثم الصنف+الموقع، ثم أولوية shipment قبل receive قبل باقي الحركات
    def ref_priority(r):
        if "shipment" in r:
            return 1
        if "receive" in r:
            return 2
        return 3

    df["_ref_priority"] = df["_ref_lower"].apply(ref_priority)
    df = df.sort_values(["Physical date", "_key", "_loc", "_ref_priority"]).reset_index(drop=True)

    batches = {}     # key = item|variant|loc -> list of batch dicts
    transfers = {}   # doc number -> queue of shipped pieces waiting to be received
    pending = {}      # كميات صرف عجزت تتغطى من رصيد موجود (احتمال خطأ ترتيب تاريخي)

    def ensure(key):
        batches.setdefault(key, [])
        pending.setdefault(key, 0)

    def fifo_deduct(key, qty, doc, is_transfer=False, loc="", trans_date=None):
        arr = batches[key]
        for b in arr:
            if qty <= 0:
                break
            if b["remaining"] <= 0:
                continue
            d = min(b["remaining"], qty)
            b["remaining"] -= d
            qty -= d
            if is_transfer:
                transfers.setdefault(doc, []).append(
                    {"qty": d, "date": b["date"], "unit_cost": b["unit_cost"], "from_loc": loc, "transfer_date": trans_date}
                )
            else:
                b["consumed"] += d
        if qty > 0:
            pending[key] += qty  # صرف من غير رصيد كافي = مؤشر مشكلة ترتيب/بيانات

    def settle(key):
        p = pending.get(key, 0)
        if p <= 0:
            return
        for b in batches[key]:
            if p <= 0:
                break
            if b["remaining"] <= 0:
                continue
            d = min(b["remaining"], p)
            b["remaining"] -= d
            p -= d
            b["adjusted"] = True
        pending[key] = p

    for _, r in df.iterrows():
        item, variant, loc = str(r["Item number"]).strip(), str(r["Variant number"]).strip(), r["_loc"]
        qty = float(r["Quantity"]) if pd.notna(r["Quantity"]) else 0
        cost = float(r["Physical cost amount"]) if pd.notna(r.get("Physical cost amount")) else 0
        ref, doc, date = r["_ref"], str(r["Number"]).strip(), r["Physical date"]
        key = f"{item}|{variant}|{loc}"
        ensure(key)
        transfers.setdefault(doc, [])
        unit_cost = cost / qty if qty > 0 else 0

        if ref == "Transfer order shipment" and qty < 0:
            fifo_deduct(key, abs(qty), doc, is_transfer=True, loc=loc, trans_date=date)

        elif ref == "Transfer order receive" and qty > 0:
            q, src = qty, ""
            while q > 0 and transfers[doc]:
                s = transfers[doc][0]
                src = s["from_loc"]
                t = min(q, s["qty"])
                batches[key].append({
                    "qty": t, "remaining": t, "consumed": 0, "date": s["date"],
                    "unit_cost": s["unit_cost"], "transfer_from": src,
                    "transfer_id": doc, "transfer_date": s["transfer_date"], "adjusted": False,
                })
                s["qty"] -= t
                q -= t
                if s["qty"] <= 0:
                    transfers[doc].pop(0)
            if q > 0:
                batches[key].append({
                    "qty": q, "remaining": q, "consumed": 0, "date": date, "unit_cost": unit_cost,
                    "transfer_from": src, "transfer_id": doc, "transfer_date": date, "adjusted": False,
                })
            settle(key)

        elif qty > 0:
            batches[key].append({
                "qty": qty, "remaining": qty, "consumed": 0, "date": date,
                "unit_cost": unit_cost, "transfer_from": "", "transfer_id": "",
                "transfer_date": None, "adjusted": False,
            })
            settle(key)

        elif qty < 0:
            fifo_deduct(key, abs(qty), doc)

    # ---- بناء تقرير الدفعات النهائي ----
    meta = (
        df.sort_values("Physical date")
        .groupby(["Item number", "Variant number"])
        .agg({
            "Description": "first", "Unit": "first",
            "Dimension 1": "first", "Dimension 2": "first", "Dimension 3": "first",
            "Dimension 4": "first", "Dimension 5": "first",
        }).to_dict("index")
    )

    out_rows = []
    for key, blist in batches.items():
        item, variant, loc = key.split("|")
        site = loc.split("-")[0]
        m = meta.get((item, variant), {})
        for b in blist:
            if abs(b["remaining"]) < 0.01:
                continue
            ratio = (b["consumed"] / b["qty"]) if b["qty"] else 0
            age = (today - b["date"]).days
            status = get_status(age, ratio)

            t_date = b["transfer_date"] or b["date"]
            t_age = (today - pd.Timestamp(t_date)).days
            t_status = get_status(t_age, ratio)

            if status in ("💣 Idle Stock", "💣 Dead Batch") and b["transfer_date"] is not None:
                status = "💣 Reallocated Idle"
            if status == "💣 Reallocated Idle" and t_status in ("💣 Idle Stock", "💣 Dead Batch"):
                t_status = "💣 Reallocated Idle (Still Idle)"

            out_rows.append({
                "Item": item, "Variant": variant, "Site": site, "Location": loc,
                "Batch Qty": round(b["qty"], 2), "Remaining Qty": round(b["remaining"], 2),
                "Unit Cost": round(b["unit_cost"], 2),
                "Remaining Value": round(b["remaining"] * b["unit_cost"], 2),
                "Batch Value": round(b["qty"] * b["unit_cost"], 2),
                "Consumption %": round(ratio * 100, 2),
                "Batch Date": b["date"].date() if pd.notna(b["date"]) else None,
                "Age (Days)": age, "Status": status,
                "Transfer Date": pd.Timestamp(t_date).date() if t_date is not None else None,
                "Age (Transfer Days)": t_age, "Status (Transfer)": t_status,
                "Transfer Info": f"🔁 From {b['transfer_from']}" if b.get("transfer_from") else "",
                "Adjusted": "✔ Covered Previous Negative" if b.get("adjusted") else "",
                "Description": m.get("Description", ""), "Unit": m.get("Unit", ""),
                "Dimension 1": m.get("Dimension 1", ""), "Dimension 2": m.get("Dimension 2", ""),
                "Dimension 3": m.get("Dimension 3", ""), "Dimension 4": m.get("Dimension 4", ""),
                "Dimension 5": m.get("Dimension 5", ""),
            })

    return pd.DataFrame(out_rows), pending


def build_value_by_site_status(batch_df: pd.DataFrame) -> pd.DataFrame:
    """تقرير: إجمالي قيمة المخزون حسب المشروع (Site) وحالة الدفعة (Status)"""
    df = batch_df.copy()
    df["Site"] = df["Site"].astype(str).fillna("")
    pivot = df.pivot_table(index="Site", columns="Status", values="Remaining Value",
                            aggfunc="sum", fill_value=0)
    pivot["Grand Total"] = pivot.sum(axis=1)
    return pivot.reset_index()


def build_detailed_status_matrix(batch_df: pd.DataFrame) -> pd.DataFrame:
    """تقرير تفصيلي: لكل صنف/موقع، الكمية والقيمة مقسمة على أعمدة كل حالة"""
    df = batch_df.copy()
    group_cols = ["Site", "Item", "Variant", "Description", "Unit",
                  "Dimension 1", "Dimension 2", "Dimension 3", "Dimension 4", "Dimension 5"]
    # توحيد النوع لنص عشان نتجنب خطأ الترتيب لما تكون الأعمدة فيها خليط أرقام/نصوص/فاضي
    for c in group_cols:
        df[c] = df[c].astype(str).replace({"nan": "", "None": ""}).fillna("")

    qty_pivot = df.pivot_table(index=group_cols, columns="Status", values="Remaining Qty",
                                aggfunc="sum", fill_value=0)
    val_pivot = df.pivot_table(index=group_cols, columns="Status", values="Remaining Value",
                                aggfunc="sum", fill_value=0)
    qty_pivot.columns = [f"{c} - Qty" for c in qty_pivot.columns]
    val_pivot.columns = [f"{c} - Value" for c in val_pivot.columns]
    merged = qty_pivot.join(val_pivot).reset_index()

    totals = df.groupby(group_cols)[["Remaining Qty", "Remaining Value"]].sum().reset_index()
    merged = merged.merge(totals, on=group_cols, how="left")
    merged = merged.rename(columns={"Remaining Qty": "Total Qty", "Remaining Value": "Total Value"})
    return merged
