from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
import html
import json
import os
import re
import sys
import traceback
import xml.etree.ElementTree as ET
import zipfile


PORT = int(os.environ.get("PORT", "8765"))
HOST = os.environ.get("HOST", "127.0.0.1")
NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_PACKAGE_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def col_to_num(col):
    total = 0
    for ch in col:
        total = total * 26 + ord(ch.upper()) - 64
    return total


def split_cell_ref(ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    return match.group(1), int(match.group(2))


def excel_date(value):
    if value is None or value == "":
        return None
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return value
    date = datetime(1899, 12, 30) + timedelta(days=serial)
    return date.date().isoformat()


def number_or_none(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        n = float(value)
        return int(n) if n.is_integer() else n
    except (TypeError, ValueError):
        return None


def boolish(value):
    if value in (None, ""):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip().lower() in {"1", "si", "sì", "yes", "true", "x", "ok"}


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def find_col(headers, candidates, fallback=None):
    normalized = {col: normalize_header(value) for col, value in headers.items()}
    wanted = [normalize_header(candidate) for candidate in candidates]
    for col, header in normalized.items():
        if header in wanted:
            return col
    for col, header in normalized.items():
        if any(candidate and candidate in header for candidate in wanted):
            return col
    return fallback


def find_header_row(rows):
    rows_found = find_header_rows(rows)
    return rows_found[0] if rows_found else None


def find_header_rows(rows):
    found = []
    for row_idx in sorted(rows.keys()):
        headers = rows[row_idx]
        col_in = find_col(headers, ["Data IN", "Check in", "Check-in"])
        col_out = find_col(headers, ["Data OUT", "Check out", "Check-out"])
        if col_in and col_out:
            found.append(row_idx)
    return found


def weekday_it(day):
    names = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
    return names[day.weekday()]


def cleaning_date_for(check_out, next_checkin=None, avoid_sunday=False):
    checkout_day = parse_iso_date(check_out)
    next_day = parse_iso_date(next_checkin)
    if not checkout_day:
        return check_out, ""
    if avoid_sunday and checkout_day.weekday() == 6 and next_day != checkout_day:
        return (checkout_day + timedelta(days=1)).isoformat(), "Domenica evitata: pulire lunedi"
    return check_out, ""


def form_number(form, name, default=0):
    value = number_or_none(form.getvalue(name))
    return default if value is None else value


def simple_formula_value(formula):
    if not formula:
        return None
    expr = formula.lstrip("=")
    if not re.fullmatch(r"[0-9\.\+\-\*/\(\) ]+", expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return None


class XlsxReader:
    def __init__(self, raw):
        self.zip = zipfile.ZipFile(BytesIO(raw))
        self.shared_strings = self._shared_strings()
        self.date_style_ids = self._date_style_ids()
        self.sheets = self._sheets()

    def _xml(self, name):
        return ET.fromstring(self.zip.read(name))

    def _shared_strings(self):
        try:
            root = self._xml("xl/sharedStrings.xml")
        except KeyError:
            return []
        values = []
        for item in root.findall(f"{NS_MAIN}si"):
            values.append("".join(t.text or "" for t in item.iter(f"{NS_MAIN}t")))
        return values

    def _date_style_ids(self):
        try:
            root = self._xml("xl/styles.xml")
        except KeyError:
            return set()
        custom_date_formats = set()
        date_format_ids = {
            14, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57,
        }
        num_fmts = root.find(f"{NS_MAIN}numFmts")
        if num_fmts is not None:
            for fmt in num_fmts.findall(f"{NS_MAIN}numFmt"):
                fmt_id = int(fmt.attrib.get("numFmtId", 0))
                code = fmt.attrib.get("formatCode", "").lower()
                if any(token in code for token in ("yy", "dd", "mm")):
                    custom_date_formats.add(fmt_id)
        style_ids = set()
        cell_xfs = root.find(f"{NS_MAIN}cellXfs")
        if cell_xfs is not None:
            for idx, xf in enumerate(cell_xfs.findall(f"{NS_MAIN}xf")):
                fmt_id = int(xf.attrib.get("numFmtId", 0))
                if fmt_id in date_format_ids or fmt_id in custom_date_formats:
                    style_ids.add(str(idx))
        return style_ids

    def _sheets(self):
        workbook = self._xml("xl/workbook.xml")
        rels = self._xml("xl/_rels/workbook.xml.rels")
        rel_map = {}
        for rel in rels.findall(f"{NS_PACKAGE_REL}Relationship"):
            rel_map[rel.attrib["Id"]] = rel.attrib["Target"]
        sheets = []
        for sheet in workbook.findall(f".//{NS_MAIN}sheet"):
            rel_id = sheet.attrib[f"{NS_REL}id"]
            target = rel_map[rel_id]
            path = "xl/" + target.lstrip("/")
            sheets.append({"name": sheet.attrib["name"], "path": path})
        return sheets

    def cell_value(self, cell):
        cell_type = cell.attrib.get("t")
        style_id = cell.attrib.get("s")
        formula = cell.find(f"{NS_MAIN}f")
        value = cell.find(f"{NS_MAIN}v")
        text = value.text if value is not None else None
        if text is None and formula is not None:
            return simple_formula_value(formula.text)
        if cell_type == "s":
            return self.shared_strings[int(text)] if text is not None else None
        if cell_type == "inlineStr":
            return "".join(t.text or "" for t in cell.iter(f"{NS_MAIN}t"))
        if style_id in self.date_style_ids:
            return excel_date(text)
        if text is None:
            return None
        try:
            n = float(text)
            return int(n) if n.is_integer() else n
        except ValueError:
            return text

    def rows(self, sheet_path):
        root = self._xml(sheet_path)
        rows = {}
        formulas = {}
        for cell in root.findall(f".//{NS_MAIN}c"):
            ref = cell.attrib.get("r")
            col, row = split_cell_ref(ref)
            rows.setdefault(row, {})[col] = self.cell_value(cell)
            formula = cell.find(f"{NS_MAIN}f")
            if formula is not None and formula.text:
                formulas[ref] = "=" + formula.text
        return rows, formulas


def used_bed_sets(guests, extra_bed, bed_guests):
    guests = number_or_none(guests)
    bed_guests = max(1, number_or_none(bed_guests) or 2)
    if guests is None:
        return 0
    beds = int((guests + bed_guests - 1) // bed_guests)
    return beds + (1 if extra_bed else 0)


def status_for(stock, threshold):
    if stock < 0:
        return "red"
    if stock <= threshold:
        return "yellow"
    return "green"


def delivery_due_date(critical_date, today=None, notice_days=2):
    if not critical_date:
        return None
    today = today or datetime.now().date().isoformat()
    try:
        day = datetime.fromisoformat(critical_date).date()
        today_date = datetime.fromisoformat(today).date()
    except ValueError:
        return None
    if day.day > 15:
        due = day.replace(day=15)
    elif day.day > 1:
        due = day.replace(day=1)
    else:
        previous_month = day.replace(day=1) - timedelta(days=1)
        due = previous_month.replace(day=15)
    minimum_notice = today_date + timedelta(days=notice_days)
    return max(due, minimum_notice).isoformat()


def add_days(date_text, days):
    if not date_text:
        return None
    try:
        return (datetime.fromisoformat(date_text).date() + timedelta(days=days)).isoformat()
    except ValueError:
        return None


def next_calendar_delivery(today, notice_days=2):
    minimum = datetime.fromisoformat(today).date() + timedelta(days=notice_days)
    candidates = [
        minimum.replace(day=15) if minimum.day <= 15 else None,
        minimum.replace(day=1),
    ]
    if minimum.day > 15:
        next_month_seed = (minimum.replace(day=28) + timedelta(days=4)).replace(day=1)
        candidates = [next_month_seed, next_month_seed.replace(day=15)]
    else:
        candidates = [minimum.replace(day=15)]
    future = [day for day in candidates if day and day >= minimum]
    return min(future).isoformat()


def latest_delivery_date(deliveries, apartment):
    dates = []
    for delivery in deliveries:
        target = str(delivery.get("apartment") or "").strip()
        if target and target != apartment:
            continue
        date_text = delivery.get("date")
        if date_text:
            dates.append(date_text)
    return max(dates) if dates else None


def first_delivery_date(deliveries, apartment):
    dates = []
    for delivery in deliveries:
        target = str(delivery.get("apartment") or "").strip()
        if target and target != apartment:
            continue
        date_text = delivery.get("date")
        if date_text:
            dates.append(date_text)
    return min(dates) if dates else None


def label_from_sheet_name(sheet_name):
    text = str(sheet_name or "").strip()
    label = re.sub(r"^(?:apt|app|id)?\s*[\-_ ]*\d+[A-Za-z]?\s*[\-_ ]+", "", text, flags=re.IGNORECASE)
    label = label.replace("_", " ").replace("-", " ")
    label = re.sub(r"\s+", " ", label).strip()
    return label or text


def is_apartment_sheet(sheet_name):
    return bool(re.match(r"^ID00\d+", str(sheet_name or "").strip(), flags=re.IGNORECASE))


def apartment_meta(settings, sheet_name):
    registry = settings.get("apartments") or []
    for item in registry:
        item_id = str(item.get("id") or "").strip()
        if item_id == sheet_name:
            label = str(item.get("label") or "").strip()
            address = str(item.get("address") or "").strip()
            def item_num(key, default):
                value = number_or_none(item.get(key))
                return default if value is None else value
            return {
                "id": sheet_name,
                "label": label or sheet_name,
                "address": address,
                "useLaundry": boolish(item.get("useLaundry")),
                "avoidSundayCleaning": boolish(item.get("avoidSundayCleaning")),
                "threshold": max(0, item_num("threshold", number_or_none(settings.get("threshold")) or 2)),
                "bedGuests": max(1, item_num("bedGuests", number_or_none(settings.get("bedGuests")) or 2)),
                "bathGuests": max(1, item_num("bathGuests", number_or_none(settings.get("bathGuests")) or 1)),
                "matsStandard": max(0, item_num("matsStandard", number_or_none(settings.get("matsStandard")) or 1)),
                "matsExtra": max(0, item_num("matsExtra", number_or_none(settings.get("matsExtra")) or 0)),
                "startBed": max(0, item_num("startBed", number_or_none(settings.get("startBed")) or 0)),
                "startBath": max(0, item_num("startBath", number_or_none(settings.get("startBath")) or 0)),
                "startMats": max(0, item_num("startMats", number_or_none(settings.get("startMats")) or 0)),
                "startDate": str(item.get("startDate") or settings.get("startDate") or "").strip(),
                "guestCol": str(item.get("guestCol") or "").strip().upper(),
                "extraBedCol": str(item.get("extraBedCol") or "").strip().upper(),
            }
    return {
        "id": sheet_name,
        "label": label_from_sheet_name(sheet_name),
        "address": "",
        "useLaundry": False,
        "avoidSundayCleaning": False,
        "threshold": max(0, number_or_none(settings.get("threshold")) or 2),
        "bedGuests": max(1, number_or_none(settings.get("bedGuests")) or 2),
        "bathGuests": max(1, number_or_none(settings.get("bathGuests")) or 1),
        "matsStandard": max(0, number_or_none(settings.get("matsStandard")) or 1),
        "matsExtra": max(0, number_or_none(settings.get("matsExtra")) or 0),
        "startBed": max(0, number_or_none(settings.get("startBed")) or 0),
        "startBath": max(0, number_or_none(settings.get("startBath")) or 0),
        "startMats": max(0, number_or_none(settings.get("startMats")) or 0),
        "startDate": str(settings.get("startDate") or "").strip(),
        "guestCol": "",
        "extraBedCol": "",
    }


def sum_consumption_between(bookings, start_date, end_date):
    total = {"bed": 0, "bath": 0, "mats": 0}
    if not start_date or not end_date:
        return total
    for booking in bookings:
        check_in = booking.get("checkIn")
        if check_in and start_date < check_in <= end_date:
            total["bed"] += number_or_none(booking["consume"].get("bed")) or 0
            total["bath"] += number_or_none(booking["consume"].get("bath")) or 0
            total["mats"] += number_or_none(booking["consume"].get("mats")) or 0
    return total


def order_needed_until(movements, cutoff_date, threshold):
    need = {"bed": 0, "bath": 0, "mats": 0}
    if not cutoff_date:
        return need
    for movement in movements:
        date_text = movement.get("checkIn")
        if not date_text or date_text > cutoff_date:
            continue
        remaining = movement.get("remaining") or {}
        need["bed"] = max(need["bed"], threshold - (number_or_none(remaining.get("bed")) or 0))
        need["bath"] = max(need["bath"], threshold - (number_or_none(remaining.get("bath")) or 0))
        need["mats"] = max(need["mats"], threshold - (number_or_none(remaining.get("mats")) or 0))
    return {key: max(0, value) for key, value in need.items()}


def order_needed_between(movements, start_date, end_date, threshold):
    need = {"bed": 0, "bath": 0, "mats": 0}
    if not start_date or not end_date:
        return need
    for movement in movements:
        date_text = movement.get("checkIn")
        if not date_text or not (start_date <= date_text <= end_date):
            continue
        remaining = movement.get("remaining") or {}
        need["bed"] = max(need["bed"], threshold - (number_or_none(remaining.get("bed")) or 0))
        need["bath"] = max(need["bath"], threshold - (number_or_none(remaining.get("bath")) or 0))
        need["mats"] = max(need["mats"], threshold - (number_or_none(remaining.get("mats")) or 0))
    return {key: max(0, value) for key, value in need.items()}


def days_between(left, right):
    try:
        return abs((datetime.fromisoformat(left).date() - datetime.fromisoformat(right).date()).days)
    except ValueError:
        return 0


def parse_iso_date(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value)).date()
    except ValueError:
        return None


def date_span(start, end):
    start_date = parse_iso_date(start)
    end_date = parse_iso_date(end)
    if not start_date or not end_date or end_date <= start_date:
        return []
    days = []
    current = start_date
    while current < end_date:
        days.append(current.isoformat())
        current += timedelta(days=1)
    return days


def money_or_none(value):
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).strip()
    cleaned = re.sub(r"[^\d,\.\-]", "", cleaned)
    if not cleaned:
        return None
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def pricing_settings(settings, apartment_id):
    selected = {}
    for row in settings.get("apartments") or []:
        if str(row.get("id") or "").strip() == apartment_id:
            selected = row
            break
    if not selected:
        for row in settings.get("pricing") or []:
            if str(row.get("id") or "").strip() == apartment_id:
                selected = row
                break

    def num(key, default):
        value = money_or_none(selected.get(key))
        return default if value is None else value

    base = num("basePrice", 90)
    weekend = num("weekendPrice", round(base * 1.18))
    min_price = num("minPrice", max(35, round(base * 0.72)))
    max_price = num("maxPrice", round(max(base, weekend) * 1.75))
    return {
        "basePrice": base,
        "weekendPrice": weekend,
        "cleaningFee": num("cleaningFee", 0),
        "extraFee": num("extraFee", 0),
        "minPrice": min_price,
        "maxPrice": max(max_price, min_price),
        "notes": str(selected.get("notes") or "").strip(),
    }


def round_price(value):
    return int(round(value / 5.0) * 5)


def build_pricing_calendar(apartment_id, label, bookings, price_config, today):
    today_date = datetime.fromisoformat(today).date()
    horizon = 45
    occupied = {}
    current_prices = {}
    booking_starts = []
    past_revenue = 0
    past_nights = 0
    future_occupied = 0

    for booking in bookings:
        check_in = booking.get("checkIn")
        check_out = booking.get("checkOut")
        nights = len(date_span(check_in, check_out))
        if nights <= 0:
            continue
        start_date = parse_iso_date(check_in)
        if start_date:
            booking_starts.append(start_date)
        amount = money_or_none(booking.get("amount"))
        if amount is not None:
            past_revenue += amount
            past_nights += nights
        nightly_current = money_or_none(booking.get("currentPrice"))
        if nightly_current is None and amount is not None:
            nightly_current = amount / nights
        for day in date_span(check_in, check_out):
            occupied[day] = booking
            if nightly_current is not None:
                current_prices[day] = round_price(nightly_current)
            if today <= day < (today_date + timedelta(days=30)).isoformat():
                future_occupied += 1

    adr = round(past_revenue / past_nights, 2) if past_nights else None
    occupancy_30 = round((future_occupied / 30) * 100, 1)
    rows = []
    available_count = 0
    total_suggested = 0

    for offset in range(horizon):
        day = today_date + timedelta(days=offset)
        day_text = day.isoformat()
        booking = occupied.get(day_text)
        weekend = day.weekday() in (4, 5)
        base = price_config["weekendPrice"] if weekend else price_config["basePrice"]
        reasons = []
        multiplier = 1.0

        if weekend:
            reasons.append("weekend")
        if occupancy_30 >= 75:
            multiplier += 0.12
            reasons.append("occupazione alta")
        elif occupancy_30 <= 35:
            multiplier -= 0.10
            reasons.append("occupazione bassa")

        days_to_arrival = offset
        if days_to_arrival <= 3:
            multiplier -= 0.18
            reasons.append("last minute")
        elif days_to_arrival <= 7:
            multiplier -= 0.10
            reasons.append("entro 7 giorni")
        elif days_to_arrival >= 30:
            multiplier += 0.08
            reasons.append("anticipo")

        previous_booking = max((d for d in booking_starts if d < day), default=None)
        next_booking = min((d for d in booking_starts if d > day), default=None)
        gap_size = None
        if previous_booking and next_booking:
            gap_size = (next_booking - previous_booking).days
        if gap_size and gap_size <= 4 and not booking:
            multiplier -= 0.12
            reasons.append("buco breve")

        suggested = round_price(max(price_config["minPrice"], min(price_config["maxPrice"], base * multiplier)))
        if booking:
            action = "occupato"
            final_price = ""
        else:
            available_count += 1
            total_suggested += suggested
            final_price = suggested
            if days_to_arrival <= 7:
                action = "spingere"
            elif occupancy_30 >= 75:
                action = "proteggere prezzo"
            else:
                action = "standard"

        rows.append({
            "date": day_text,
            "weekday": weekday_it(day),
            "occupied": bool(booking),
            "guestName": booking.get("guestName") if booking else "",
            "currentPrice": current_prices.get(day_text),
            "basePrice": base,
            "finalPrice": final_price,
            "action": action,
            "reason": ", ".join(reasons) if reasons else "base",
        })

    return {
        "apartment": apartment_id,
        "apartmentLabel": label,
        "settings": price_config,
        "adr": adr,
        "occupancy30": occupancy_30,
        "available45": available_count,
        "averageSuggested": round(total_suggested / available_count, 2) if available_count else None,
        "calendar": rows,
    }


def checkout_dates_for(booking_events):
    return sorted({
        event["checkOut"]
        for event in booking_events
        if event.get("kind") == "booking" and event.get("checkOut")
    })


def align_delivery_to_checkout(date_text, booking_events, earliest_date=None, latest_date=None):
    if not date_text:
        return None
    earliest = earliest_date or date_text
    bookings = [event for event in booking_events if event.get("kind") == "booking"]
    aligned = max(date_text, earliest)
    while True:
        covering = next((
            event for event in bookings
            if event.get("checkIn") and event.get("checkOut") and event["checkIn"] < aligned < event["checkOut"]
        ), None)
        if not covering:
            break
        aligned = covering["checkOut"]
    candidates = [
        day for day in checkout_dates_for(booking_events)
        if day >= aligned and (not latest_date or day <= latest_date)
    ]
    if candidates:
        return candidates[0]
    return aligned


def parse_workbook(raw, settings=None, deliveries=None):
    settings = settings or {}
    deliveries = deliveries or []
    reader = XlsxReader(raw)
    apartment_sheets = [sheet for sheet in reader.sheets if is_apartment_sheet(sheet["name"])]
    ignored_sheets = [sheet["name"] for sheet in reader.sheets if not is_apartment_sheet(sheet["name"])]
    parse_status = {
        sheet["name"]: {"ok": False, "message": "Non ancora letto"}
        for sheet in apartment_sheets
    }
    apartments = []
    all_critical = []
    all_bookings = []
    all_reservations = []
    delivery_schedule = []
    pricing_results = []
    cleaning_schedule = []
    total_bookings = 0
    today = datetime.now().date().isoformat()
    default_threshold = max(0, int(settings.get("threshold", 2)))
    for sheet in apartment_sheets:
        meta = apartment_meta(settings, sheet["name"])
        use_laundry = bool(meta.get("useLaundry"))
        avoid_sunday_cleaning = bool(meta.get("avoidSundayCleaning"))
        threshold = int(meta["threshold"])
        bed_guests = int(meta["bedGuests"])
        bath_guests = int(meta["bathGuests"])
        mats_standard = int(meta["matsStandard"])
        mats_extra = int(meta["matsExtra"])
        initial_stock = {
            "bed": int(meta["startBed"]),
            "bath": int(meta["startBath"]),
            "mats": int(meta["startMats"]),
        }
        rows, formulas = reader.rows(sheet["path"])
        header_rows = find_header_rows(rows)
        if not header_rows:
            parse_status[sheet["name"]] = {
                "ok": False,
                "message": "Non trovo intestazioni con Data IN e Data OUT",
            }
            continue

        base_events = []
        pricing_events = []
        start_from = meta.get("startDate") or today
        configured_guest_col = meta.get("guestCol") or ""
        configured_extra_col = meta.get("extraBedCol") or ""
        parsed_tables = []
        missing_guests_tables = []
        detected_cols = {"guests": "", "extra": "", "amount": "", "current": ""}

        for idx, header_row in enumerate(header_rows):
            headers = rows[header_row]
            col_ota = find_col(headers, ["OTA"], "B")
            col_booking = find_col(headers, ["Prenotazione", "Codice prenotazione"], "C")
            col_in = find_col(headers, ["Data IN", "Check in", "Check-in"])
            col_out = find_col(headers, ["Data OUT", "Check out", "Check-out"])
            col_name = find_col(headers, ["Nome", "Nome ospite", "Guest"])
            col_amount = find_col(headers, [
                "Totale",
                "Importo",
                "Totale prenotazione",
                "Revenue",
                "Incasso",
                "Lordo",
                "Netto",
                "Pagamento",
            ])
            col_current_price = find_col(headers, [
                "Prezzo notte",
                "Prezzo attuale",
                "Tariffa notte",
                "Tariffa",
                "Daily price",
                "Nightly rate",
            ])
            col_guests = configured_guest_col or find_col(headers, [
                "N° Ospiti",
                "N Ospiti",
                "Nr Ospiti",
                "Numero ospiti",
                "Ospiti",
                "Persone",
                "Pax",
                "Guest number",
                "Guests",
            ])
            col_extra_bed = configured_extra_col or find_col(headers, [
                "Letto extra",
                "Eccezione letto",
                "Extra letto",
                "Letti extra",
                "Letto aggiuntivo",
                "Extra bed",
            ])
            detected_cols = {
                "guests": detected_cols["guests"] or col_guests or "",
                "extra": detected_cols["extra"] or col_extra_bed or "",
                "amount": detected_cols["amount"] or col_amount or "",
                "current": detected_cols["current"] or col_current_price or "",
            }
            next_header = header_rows[idx + 1] if idx + 1 < len(header_rows) else None
            table_events = 0
            table_pricing_events = 0
            candidate_rows = sorted(r for r in rows.keys() if r > header_row and (next_header is None or r < next_header))
            for row_idx in candidate_rows:
                row = rows[row_idx]
                booking_id = row.get(col_booking)
                check_in = row.get(col_in)
                check_out = row.get(col_out)
                has_dates = bool(check_in and check_out)
                if not has_dates:
                    continue
                if check_in < start_from:
                    continue
                pricing_events.append({
                    "kind": "booking",
                    "date": check_in,
                    "row": row_idx,
                    "apartment": sheet["name"],
                    "apartmentLabel": meta["label"],
                    "deliveryAddress": meta["address"],
                    "ota": row.get(col_ota) or "",
                    "booking": booking_id or "",
                    "guestName": (row.get(col_name) or "").strip() if isinstance(row.get(col_name), str) else row.get(col_name) or "",
                    "checkIn": check_in,
                    "checkOut": check_out,
                    "guests": number_or_none(row.get(col_guests)) if col_guests else None,
                    "extraBed": boolish(row.get(col_extra_bed)) if col_extra_bed else False,
                    "amount": money_or_none(row.get(col_amount)) if col_amount else None,
                    "currentPrice": money_or_none(row.get(col_current_price)) if col_current_price else None,
                })
                table_pricing_events += 1
                if not col_guests:
                    continue
                guests = number_or_none(row.get(col_guests))
                if not guests:
                    continue
                extra_bed = boolish(row.get(col_extra_bed)) if col_extra_bed else False
                consume = {
                    "bed": used_bed_sets(guests, extra_bed, bed_guests),
                    "bath": int((guests + bath_guests - 1) // bath_guests),
                    "mats": mats_standard + (mats_extra if extra_bed else 0),
                }
                base_events.append({
                    "kind": "booking",
                    "date": check_in,
                    "row": row_idx,
                    "apartment": sheet["name"],
                    "apartmentLabel": meta["label"],
                    "deliveryAddress": meta["address"],
                    "ota": row.get(col_ota) or "",
                    "booking": booking_id or "",
                    "guestName": (row.get(col_name) or "").strip() if isinstance(row.get(col_name), str) else row.get(col_name) or "",
                    "checkIn": check_in,
                    "checkOut": check_out,
                    "guests": guests,
                    "amount": money_or_none(row.get(col_amount)) if col_amount else None,
                    "currentPrice": money_or_none(row.get(col_current_price)) if col_current_price else None,
                    "extraBed": extra_bed,
                    "consume": consume,
                })
                table_events += 1
            if col_in and col_out and not col_guests and table_pricing_events:
                missing_guests_tables.append(str(header_row))
            parsed_tables.append({"row": header_row, "guests": col_guests, "events": table_events, "pricingEvents": table_pricing_events})

        if not parsed_tables or not any(item["pricingEvents"] for item in parsed_tables):
            message = "Manca colonna Ospiti nelle tabelle con Data IN e Data OUT"
            if missing_guests_tables:
                message += " (righe intestazione: " + ", ".join(missing_guests_tables) + ")"
            parse_status[sheet["name"]] = {"ok": False, "message": message}
            continue

        parse_status[sheet["name"]] = {
            "ok": True,
            "message": "Tabelle lette: " + ", ".join(
                f"riga {item['row']} ospiti {item['guests'] or 'mancante'} ({item['pricingEvents']} date, {item['events']} bianch.)"
                for item in parsed_tables
            ),
        }

        price_config = pricing_settings(settings, sheet["name"])
        pricing_results.append(build_pricing_calendar(sheet["name"], meta["label"], pricing_events, price_config, today))
        all_reservations.extend(pricing_events)
        sorted_pricing_events = sorted(pricing_events, key=lambda event: (event.get("checkIn") or "", event.get("checkOut") or ""))
        for idx, event in enumerate(sorted_pricing_events):
            check_out = event.get("checkOut")
            if not check_out or check_out < today:
                continue
            future_events = [
                next_event
                for next_event in sorted_pricing_events[idx + 1:]
                if next_event.get("checkIn") and next_event.get("checkIn") >= check_out
            ]
            next_event = future_events[0] if future_events else None
            next_checkin = next_event.get("checkIn") if next_event else None
            next_guests = next_event.get("guests") if next_event else None
            next_extra_bed = bool(next_event.get("extraBed")) if next_event else False
            beds_to_prepare = used_bed_sets(next_guests, next_extra_bed, bed_guests) if next_guests else None
            intervention_date, cleaning_rule_note = cleaning_date_for(check_out, next_checkin, avoid_sunday=avoid_sunday_cleaning)
            if next_checkin == check_out:
                priority = "Alta"
                note = "Check-in stesso giorno"
            elif next_checkin and next_checkin <= add_days(check_out, 1):
                priority = "Media"
                note = "Check-in il giorno dopo"
            else:
                priority = "Normale"
                note = ""
            if cleaning_rule_note:
                note = f"{note}. {cleaning_rule_note}" if note else cleaning_rule_note
            cleaning_schedule.append({
                "date": intervention_date,
                "apartment": sheet["name"],
                "apartmentLabel": meta["label"],
                "checkout": check_out,
                "checkoutTime": "10:30",
                "nextCheckin": next_checkin,
                "nextCheckinTime": "16:00" if next_checkin else "",
                "checkin": event.get("checkIn"),
                "guestName": event.get("guestName") or "",
                "outgoingGuests": event.get("guests"),
                "nextGuests": next_guests,
                "bedsToPrepare": beds_to_prepare,
                "ota": event.get("ota") or "",
                "priority": priority,
                "avoidSundayCleaning": avoid_sunday_cleaning,
                "note": note,
            })
        if not use_laundry:
            continue
        if not base_events:
            parse_status[sheet["name"]] = {
                "ok": False,
                "message": "Biancheria attiva, ma non trovo righe con colonna Ospiti valorizzata",
            }
            continue

        for delivery in deliveries:
            target = str(delivery.get("apartment") or "").strip()
            applies = not target or target == sheet["name"]
            if not applies:
                continue
            if delivery.get("date") and delivery.get("date") < start_from:
                continue
            base_events.append({
                "kind": "delivery",
                "date": delivery.get("date") or "0000-00-00",
                "apartment": sheet["name"],
                "supply": {
                    "bed": number_or_none(delivery.get("bed")) or 0,
                    "bath": number_or_none(delivery.get("bath")) or 0,
                    "mats": number_or_none(delivery.get("mats")) or 0,
                },
                "note": delivery.get("note") or "",
            })

        def replay(events, include_delivery_rows=False):
            stock = dict(initial_stock)
            movements = []
            critical_date = None
            critical_rows = []
            for event in sorted(events, key=lambda e: (e.get("date") or "", 0 if e["kind"] in {"delivery", "planned_delivery"} else 1)):
                if event["kind"] in {"delivery", "planned_delivery"}:
                    supply = event["supply"]
                    stock["bed"] += supply["bed"]
                    stock["bath"] += supply["bath"]
                    stock["mats"] += supply["mats"]
                    if include_delivery_rows:
                        row_statuses = {
                            "bed": status_for(stock["bed"], threshold),
                            "bath": status_for(stock["bath"], threshold),
                            "mats": status_for(stock["mats"], threshold),
                        }
                        state = "red" if "red" in row_statuses.values() else "yellow" if "yellow" in row_statuses.values() else "green"
                        movements.append({
                            "type": event["kind"],
                            "row": "",
                            "apartment": sheet["name"],
                            "apartmentLabel": meta["label"],
                            "deliveryAddress": meta["address"],
                            "ota": "",
                            "booking": "",
                            "guestName": event.get("note") or ("Consegna programmata" if event["kind"] == "planned_delivery" else "Consegna"),
                            "checkIn": event["date"],
                            "checkOut": event["date"],
                            "guests": "",
                            "extraBed": False,
                            "consume": {"bed": 0, "bath": 0, "mats": 0},
                            "supply": dict(supply),
                            "remaining": dict(stock),
                            "threshold": threshold,
                            "need": {
                                "bed": max(0, threshold - stock["bed"]),
                                "bath": max(0, threshold - stock["bath"]),
                                "mats": max(0, threshold - stock["mats"]),
                            },
                            "status": state,
                            "isPast": bool(event["date"] and event["date"] < today),
                        })
                    continue

                consume = event["consume"]
                stock["bed"] -= consume["bed"]
                stock["bath"] -= consume["bath"]
                stock["mats"] -= consume["mats"]
                row_statuses = {
                    "bed": status_for(stock["bed"], threshold),
                    "bath": status_for(stock["bath"], threshold),
                    "mats": status_for(stock["mats"], threshold),
                }
                state = "red" if "red" in row_statuses.values() else "yellow" if "yellow" in row_statuses.values() else "green"
                is_past = bool(event["checkOut"] and event["checkOut"] < today)
                if state != "green" and critical_date is None and not is_past:
                    critical_date = event["checkIn"]
                movement = {
                    "type": "booking",
                    "row": event["row"],
                    "apartment": sheet["name"],
                    "apartmentLabel": meta["label"],
                    "deliveryAddress": meta["address"],
                    "ota": event["ota"],
                    "booking": event["booking"],
                    "guestName": event["guestName"],
                    "checkIn": event["checkIn"],
                    "checkOut": event["checkOut"],
                    "guests": event["guests"],
                    "amount": event.get("amount"),
                    "currentPrice": event.get("currentPrice"),
                    "extraBed": event["extraBed"],
                    "consume": consume,
                    "supply": {"bed": 0, "bath": 0, "mats": 0},
                    "remaining": dict(stock),
                    "threshold": threshold,
                    "need": {
                        "bed": max(0, threshold - stock["bed"]),
                        "bath": max(0, threshold - stock["bath"]),
                        "mats": max(0, threshold - stock["mats"]),
                    },
                    "status": state,
                    "isPast": is_past,
                }
                movements.append(movement)
                if state != "green":
                    critical_rows.append(movement)
            return {
                "stock": stock,
                "movements": movements,
                "criticalDate": critical_date,
                "criticalRows": critical_rows,
            }

        preliminary = replay(base_events, include_delivery_rows=False)
        last_delivery = latest_delivery_date(deliveries, sheet["name"])
        minimum_delivery = add_days(today, 2)
        theoretical_delivery = add_days(last_delivery, 15) if last_delivery else next_calendar_delivery(today, notice_days=2)
        if theoretical_delivery and minimum_delivery and theoretical_delivery < minimum_delivery:
            theoretical_delivery = minimum_delivery
        scheduled_delivery = align_delivery_to_checkout(theoretical_delivery, base_events, earliest_date=minimum_delivery)
        first_window_end = add_days(scheduled_delivery, 15)
        first_order = order_needed_between(preliminary["movements"], scheduled_delivery, first_window_end, threshold)
        has_first_order = any(number_or_none(value) for value in first_order.values())
        critical_before_scheduled = None
        anticipation_start = last_delivery or today
        if has_first_order and scheduled_delivery:
            for movement in preliminary["movements"]:
                movement_date = movement.get("checkIn")
                if (
                    movement_date
                    and anticipation_start < movement_date < scheduled_delivery
                    and any(number_or_none(value) for value in (movement.get("need") or {}).values())
                    and not movement.get("isPast")
                ):
                    critical_before_scheduled = movement_date
                    break
        deliver_by = delivery_due_date(critical_before_scheduled, today=today, notice_days=2)
        if deliver_by:
            deliver_by = align_delivery_to_checkout(deliver_by, base_events, earliest_date=minimum_delivery, latest_date=critical_before_scheduled)
        suggested_delivery = deliver_by or (scheduled_delivery if has_first_order else None)
        should_anticipate = bool(deliver_by and scheduled_delivery and deliver_by < scheduled_delivery)
        second_theoretical = first_window_end
        second_delivery = align_delivery_to_checkout(second_theoretical, base_events, earliest_date=second_theoretical)
        second_window_end = add_days(second_delivery, 15) if second_delivery else None
        second_order = order_needed_between(preliminary["movements"], second_delivery, second_window_end, threshold)
        schedule_rows = []
        if has_first_order:
            schedule_rows.append({
                "apartment": sheet["name"],
                "apartmentLabel": meta["label"],
                "deliveryAddress": meta["address"],
                "sequence": 1,
                "lastDeliveryDate": last_delivery,
                "scheduledDeliveryDate": scheduled_delivery,
                "suggestedDeliveryDate": suggested_delivery,
                "deliverByDate": deliver_by,
                "shouldAnticipate": should_anticipate,
                "criticalDate": critical_before_scheduled,
                "order": dict(first_order),
                "threshold": threshold,
                "status": "red" if any(value > threshold for value in first_order.values()) else "yellow",
                "note": f"Anticipare: sotto soglia il {critical_before_scheduled}" if should_anticipate else "Ordinaria",
            })
        if any(number_or_none(value) for value in second_order.values()):
            visible_sequence = len(schedule_rows) + 1
            schedule_rows.append({
                "apartment": sheet["name"],
                "apartmentLabel": meta["label"],
                "deliveryAddress": meta["address"],
                "sequence": visible_sequence,
                "lastDeliveryDate": (suggested_delivery or scheduled_delivery) if schedule_rows else last_delivery,
                "scheduledDeliveryDate": second_delivery,
                "suggestedDeliveryDate": second_delivery,
                "deliverByDate": None,
                "shouldAnticipate": False,
                "order": second_order,
                "threshold": threshold,
                "status": "green",
                "note": "Previsione 15 giorni",
            })
        delivery_schedule.extend(schedule_rows)

        planned_events = []
        for row in schedule_rows:
            if not (number_or_none(row["order"]["bed"]) or number_or_none(row["order"]["bath"]) or number_or_none(row["order"]["mats"])):
                continue
            planned_events.append({
                "kind": "planned_delivery",
                "date": row["suggestedDeliveryDate"],
                "apartment": sheet["name"],
                "supply": dict(row["order"]),
                "note": f"Consegna programmata {row['sequence']}",
            })

        final = replay(base_events + planned_events, include_delivery_rows=True)
        bookings = final["movements"]
        critical_date = final["criticalDate"]
        stock = final["stock"]
        critical_rows = final["criticalRows"]

        total_bookings += len([row for row in bookings if row.get("type") == "booking"])
        all_bookings.extend(bookings)
        all_critical.extend(critical_rows)
        final_statuses = {
            "bed": status_for(stock["bed"], threshold),
            "bath": status_for(stock["bath"], threshold),
            "mats": status_for(stock["mats"], threshold),
        }
        apartment_status = "red" if "red" in final_statuses.values() else "yellow" if "yellow" in final_statuses.values() else "green"
        need = {
            "bed": max(0, threshold - stock["bed"]),
            "bath": max(0, threshold - stock["bath"]),
            "mats": max(0, threshold - stock["mats"]),
        }

        apartments.append({
            "name": sheet["name"],
            "label": meta["label"],
            "address": meta["address"],
            "settings": {
                "threshold": threshold,
                "useLaundry": use_laundry,
                "avoidSundayCleaning": avoid_sunday_cleaning,
                "bedGuests": bed_guests,
                "bathGuests": bath_guests,
                "matsStandard": mats_standard,
                "matsExtra": mats_extra,
                "startBed": initial_stock["bed"],
                "startBath": initial_stock["bath"],
                "startMats": initial_stock["mats"],
                "startDate": start_from,
                "guestCol": detected_cols["guests"],
                "extraBedCol": detected_cols["extra"],
                "amountCol": detected_cols["amount"],
                "currentPriceCol": detected_cols["current"],
            },
            "initial": dict(initial_stock),
            "stock": stock,
            "need": first_order,
            "projectedNeed": need,
            "status": apartment_status,
            "firstCriticalDate": critical_date,
            "deliverByDate": delivery_due_date(critical_date, today=today, notice_days=2),
            "scheduledDeliveryDate": scheduled_delivery,
            "suggestedDeliveryDate": suggested_delivery,
            "bookings": bookings,
        })

    priced_apartments = {row["apartment"] for row in pricing_results}
    for sheet in apartment_sheets:
        if sheet["name"] in priced_apartments:
            continue
        meta = apartment_meta(settings, sheet["name"])
        price_config = pricing_settings(settings, sheet["name"])
        pricing_results.append(build_pricing_calendar(sheet["name"], meta["label"], [], price_config, today))

    return {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "workbookSheets": [
            {
                "name": sheet["name"],
                "label": label_from_sheet_name(sheet["name"]),
                "address": "",
                "parseStatus": parse_status.get(sheet["name"], {"ok": False, "message": "Non letto"}),
            }
            for sheet in apartment_sheets
        ],
        "ignoredSheets": ignored_sheets,
        "parseStatus": parse_status,
        "apartments": apartments,
        "critical": sorted(all_critical, key=lambda b: b.get("checkIn") or "")[:50],
        "bookings": sorted(all_bookings, key=lambda b: (b.get("checkIn") or "", b.get("apartment") or "", 0 if b.get("type") != "booking" else 1)),
        "reservations": sorted(all_reservations, key=lambda b: (b.get("checkIn") or "", b.get("apartmentLabel") or b.get("apartment") or "")),
        "deliverySchedule": sorted(delivery_schedule, key=lambda row: (row.get("suggestedDeliveryDate") or "", row.get("apartment") or "")),
        "pricing": sorted(pricing_results, key=lambda row: row.get("apartmentLabel") or row.get("apartment") or ""),
        "cleaningSchedule": sorted(cleaning_schedule, key=lambda row: (row.get("date") or "", row.get("apartmentLabel") or row.get("apartment") or "")),
        "summary": {
            "apartments": len(apartments),
            "bookings": total_bookings,
            "critical": len(all_critical),
            "threshold": threshold,
            "defaultThreshold": default_threshold,
            "settings": {
                "bedGuests": settings.get("bedGuests", 2),
                "bathGuests": settings.get("bathGuests", 1),
                "matsStandard": settings.get("matsStandard", 1),
                "matsExtra": settings.get("matsExtra", 0),
                "startDate": settings.get("startDate", ""),
            },
        },
    }

INDEX_HTML = r"""<!doctype html>
<html lang="it">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Gestione appartamenti</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #17212b;
      --muted: #627181;
      --line: #d8e0e7;
      --panel: #ffffff;
      --bg: #f4f7f8;
      --green: #147d55;
      --yellow: #a76500;
      --red: #b42318;
      --blue: #1e5c8a;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: var(--bg);
      color: var(--ink);
    }
    header {
      padding: 24px clamp(16px, 4vw, 44px) 16px;
      background: #fff;
      border-bottom: 1px solid var(--line);
    }
    h1 { margin: 0 0 8px; font-size: clamp(24px, 4vw, 38px); letter-spacing: 0; }
    p { margin: 0; color: var(--muted); line-height: 1.45; }
    main { padding: 22px clamp(16px, 4vw, 44px) 44px; }
    .upload {
      display: grid;
      grid-template-columns: minmax(220px, 1fr) auto;
      gap: 16px;
      padding: 16px;
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      margin-bottom: 18px;
    }
    input[type=file] {
      max-width: 100%;
      width: 100%;
      padding: 14px;
      min-height: 54px;
      border: 2px dashed var(--line);
      border-radius: 8px;
      background: #fff;
      font-size: 16px;
      font-weight: 700;
      cursor: pointer;
    }
    input[type=file]::file-selector-button {
      border: 0;
      border-radius: 6px;
      background: var(--blue);
      color: #fff;
      padding: 12px 18px;
      margin-right: 14px;
      font: inherit;
      font-weight: 800;
      cursor: pointer;
    }
    .top-actions {
      display: flex;
      justify-content: flex-end;
      gap: 10px;
      flex-wrap: wrap;
    }
    .settings {
      grid-column: 1 / -1;
      display: grid;
      grid-template-columns: repeat(4, minmax(120px, 1fr));
      gap: 12px;
      padding-top: 12px;
      border-top: 1px solid var(--line);
    }
    .advanced-settings {
      grid-column: 1 / -1;
      border-top: 1px solid var(--line);
      padding-top: 12px;
    }
    .advanced-settings summary {
      cursor: pointer;
      color: var(--ink);
      font-weight: 800;
      padding: 8px 0;
    }
    .advanced-settings .settings {
      border-top: 0;
      padding-top: 10px;
    }
    .control-strip {
      display: grid;
      grid-template-columns: minmax(220px, 360px) auto auto;
      gap: 12px;
      align-items: end;
      margin-bottom: 18px;
    }
    .config-panel {
      display: none;
      margin-bottom: 18px;
    }
    .config-panel.active {
      display: block;
    }
    .config-panel h2 {
      margin: 0;
      padding: 14px 16px;
      font-size: 18px;
      border-bottom: 1px solid var(--line);
    }
    .field {
      display: grid;
      gap: 6px;
      color: var(--muted);
      font-size: 14px;
      font-weight: 700;
    }
    .field input {
      width: 100%;
      padding: 8px;
      border: 1px solid var(--line);
      border-radius: 6px;
      font: inherit;
      color: var(--ink);
      background: #fff;
    }
    .subhead {
      grid-column: 1 / -1;
      margin: 4px 0 0;
      font-size: 13px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: .04em;
      font-weight: 800;
    }
    .delivery-actions {
      grid-column: 1 / -1;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }
    .delivery-panel {
      grid-column: 1 / -1;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: #fff;
    }
    .delivery-panel table {
      min-width: 860px;
    }
    .delivery-panel th, .delivery-panel td {
      padding: 6px 8px;
    }
    .delivery-panel input {
      width: 100%;
      min-width: 80px;
      padding: 6px;
      border: 1px solid var(--line);
      border-radius: 5px;
      font: inherit;
    }
    .delivery-panel input[type="number"] {
      min-width: 62px;
    }
    .staff-panel {
      grid-column: 1 / -1;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: #fff;
    }
    .staff-panel table {
      min-width: 940px;
    }
    .staff-panel th, .staff-panel td {
      padding: 6px 8px;
    }
    .staff-panel input {
      width: 100%;
      min-width: 80px;
      padding: 6px;
      border: 1px solid var(--line);
      border-radius: 5px;
      font: inherit;
    }
    .staff-panel input[type="checkbox"] {
      width: auto;
      min-width: 0;
    }
    .apartment-panel {
      grid-column: 1 / -1;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: #fff;
    }
    .apartment-panel table {
      min-width: 1420px;
    }
    .apartment-panel th, .apartment-panel td {
      padding: 6px 8px;
    }
    .apartment-panel input, .toolbar select {
      width: 100%;
      padding: 7px;
      border: 1px solid var(--line);
      border-radius: 5px;
      font: inherit;
      background: #fff;
      color: var(--ink);
    }
    .inline-actions {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      padding: 12px 16px;
      border-top: 1px solid var(--line);
      align-items: end;
    }
    .inline-actions .field {
      min-width: 260px;
      flex: 1;
    }
    .tiny {
      padding: 6px 9px;
      font-size: 12px;
    }
    .hidden-file {
      display: none;
    }
    .toolbar { margin-bottom: 0; }
    .ghost {
      background: #e8eef2;
      color: var(--ink);
    }
    button {
      border: 0;
      border-radius: 6px;
      background: var(--blue);
      color: white;
      padding: 10px 14px;
      font-weight: 700;
      cursor: pointer;
    }
    button:disabled { opacity: .55; cursor: wait; }
    .module-nav {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }
    .module-button {
      display: grid;
      gap: 5px;
      min-height: 78px;
      padding: 14px;
      text-align: left;
      background: #fff;
      color: var(--ink);
      border: 1px solid var(--line);
    }
    .module-button strong {
      font-size: 17px;
    }
    .module-button span {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.25;
    }
    .module-button.active {
      background: var(--blue);
      color: #fff;
      border-color: var(--blue);
    }
    .module-button.active span {
      color: rgba(255, 255, 255, .82);
    }
    .module-view {
      display: none;
    }
    .module-view.active {
      display: block;
    }
    .coming-soon {
      padding: 28px;
    }
    .coming-soon h2 {
      margin: 0 0 8px;
      padding: 0;
      border: 0;
      font-size: 22px;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }
    .kpi, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
    }
    .kpi { padding: 14px; }
    .kpi strong { display: block; font-size: 28px; }
    .kpi span { color: var(--muted); font-size: 13px; }
    .pricing-dashboard {
      margin-bottom: 18px;
    }
    .calendar-panel {
      overflow: hidden;
    }
    .calendar-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
    }
    .calendar-head h2 {
      margin: 0;
      font-size: 18px;
    }
    .calendar-controls {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .calendar-controls strong {
      min-width: 170px;
      text-align: center;
      text-transform: capitalize;
    }
    .calendar-weekdays,
    .calendar-grid {
      display: grid;
      grid-template-columns: repeat(7, minmax(0, 1fr));
    }
    .calendar-weekdays {
      background: #f9fbfc;
      color: var(--muted);
      border-bottom: 1px solid var(--line);
      font-size: 12px;
      font-weight: 900;
      text-align: center;
    }
    .calendar-weekdays span {
      padding: 10px 6px;
    }
    .calendar-day {
      min-height: 126px;
      padding: 8px;
      border-right: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      background: #fff;
    }
    .calendar-day:nth-child(7n) {
      border-right: 0;
    }
    .calendar-day.outside {
      background: #f9fbfc;
      color: var(--muted);
    }
    .calendar-day-number {
      display: block;
      margin-bottom: 7px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 900;
    }
    .calendar-booking {
      z-index: 2;
      align-self: start;
      min-width: 0;
      height: 28px;
      margin: 29px 4px 0;
      padding: 6px 9px;
      overflow: hidden;
      border: 1px solid #c6dbe8;
      border-radius: 8px;
      background: #e9f2f7;
      color: var(--blue);
      font-size: 12px;
      font-weight: 900;
      line-height: 1.15;
      white-space: nowrap;
      text-overflow: ellipsis;
    }
    .calendar-booking.continues-left {
      margin-left: 0;
      border-left: 0;
      border-top-left-radius: 0;
      border-bottom-left-radius: 0;
    }
    .calendar-booking.continues-right {
      margin-right: 0;
      border-right: 0;
      border-top-right-radius: 0;
      border-bottom-right-radius: 0;
    }
    .calendar-booking small {
      color: var(--muted);
      font-size: 11px;
      font-weight: 800;
    }
    .calendar-event small {
      color: var(--muted);
      font-weight: 800;
    }
    .pricing-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
    }
    .pricing-head h2 {
      margin: 0;
      font-size: 18px;
    }
    .pricing-head span {
      color: var(--muted);
      font-size: 13px;
      font-weight: 700;
    }
    .pricing-cards {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
      gap: 14px;
      padding: 14px;
    }
    .pricing-card {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      overflow: hidden;
    }
    .pricing-card h3 {
      margin: 0;
      padding: 12px;
      font-size: 16px;
      border-bottom: 1px solid var(--line);
    }
    .pricing-metrics {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 8px;
      padding: 12px;
      border-bottom: 1px solid var(--line);
      background: #f9fbfc;
    }
    .pricing-metrics strong {
      display: block;
      font-size: 20px;
      font-variant-numeric: tabular-nums;
    }
    .pricing-metrics span {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .pricing-fields {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      padding: 12px;
    }
    .pricing-fields .field:last-child {
      grid-column: 1 / -1;
    }
    .pricing-fields input {
      width: 100%;
      padding: 8px;
      border: 1px solid var(--line);
      border-radius: 6px;
      font: inherit;
    }
    .pricing-table {
      max-height: 420px;
      overflow: auto;
      border-top: 1px solid var(--line);
    }
    .pricing-table table {
      min-width: 720px;
    }
    .price-pill {
      display: inline-flex;
      min-width: 62px;
      justify-content: center;
      padding: 4px 8px;
      border-radius: 999px;
      background: #e9f2f7;
      color: var(--blue);
      font-weight: 800;
      font-variant-numeric: tabular-nums;
    }
    .occupied-price {
      color: var(--muted);
      font-weight: 700;
    }
    .date-chip {
      display: inline-grid;
      gap: 2px;
      min-width: 126px;
      padding: 8px 10px;
      border-radius: 8px;
      font-weight: 900;
      line-height: 1.1;
      border: 1px solid transparent;
    }
    .date-chip span {
      font-size: 12px;
      font-weight: 800;
    }
    .date-chip.cleaning {
      background: #e9f2f7;
      color: var(--blue);
      border-color: #c6dbe8;
    }
    .date-chip.checkout {
      background: #fff1cc;
      color: var(--yellow);
      border-color: #ead28f;
    }
    .date-chip.arrival {
      background: #e3f5ee;
      color: var(--green);
      border-color: #bddfce;
    }
    .date-chip span {
      color: currentColor;
      opacity: .72;
    }
    .time-pill {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: fit-content;
      min-width: 54px;
      padding: 3px 7px;
      border-radius: 999px;
      background: #fff;
      color: var(--ink);
      border: 1px solid rgba(23, 33, 43, .14);
      font-size: 12px;
      font-weight: 900;
      font-variant-numeric: tabular-nums;
    }
    .metric-chip {
      display: inline-grid;
      place-items: center;
      min-width: 58px;
      min-height: 48px;
      padding: 6px 8px;
      border-radius: 8px;
      background: #f9fbfc;
      border: 1px solid var(--line);
      color: var(--ink);
      font-weight: 900;
      font-variant-numeric: tabular-nums;
    }
    .metric-chip strong {
      font-size: 22px;
      line-height: 1;
    }
    .metric-chip span {
      color: var(--muted);
      font-size: 11px;
      font-weight: 800;
      text-transform: uppercase;
    }
    .metric-cell {
      text-align: center;
    }
    .th-chip {
      display: inline-grid;
      place-items: center;
      min-width: 104px;
      min-height: 34px;
      padding: 5px 8px;
      border-radius: 8px;
      border: 1px solid var(--line);
      font-size: 11px;
      line-height: 1.05;
      text-align: center;
      white-space: normal;
    }
    .th-chip.cleaning {
      background: #e9f2f7;
      color: var(--blue);
      border-color: #c6dbe8;
    }
    .th-chip.checkout {
      background: #fff1cc;
      color: var(--yellow);
      border-color: #ead28f;
    }
    .th-chip.arrival {
      background: #e3f5ee;
      color: var(--green);
      border-color: #bddfce;
    }
    .th-chip.metric {
      min-width: 70px;
      background: #fff;
      color: var(--ink);
    }
    .status-strip {
      display: none;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 14px;
      padding: 14px 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
    }
    .status-strip.active {
      display: flex;
    }
    .status-strip strong {
      display: block;
      font-size: clamp(22px, 3vw, 32px);
      line-height: 1;
    }
    .status-strip span {
      color: var(--muted);
      font-size: 13px;
      font-weight: 800;
    }
    .cleaning-tools {
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      gap: 10px;
      padding: 12px 16px;
      border-bottom: 1px solid var(--line);
      background: #f9fbfc;
    }
    .legend {
      display: grid;
      grid-template-columns: repeat(3, minmax(180px, 1fr));
      gap: 10px;
      color: var(--ink);
      font-size: 14px;
      font-weight: 900;
    }
    .legend span {
      display: flex;
      align-items: center;
      min-height: 42px;
      padding: 9px 11px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      box-shadow: 0 1px 0 rgba(23, 33, 43, .04);
    }
    .legend span::before {
      content: "";
      width: 8px;
      height: 8px;
      margin-right: 8px;
      border-radius: 999px;
      background: var(--blue);
      flex: 0 0 auto;
    }
    body.staff-mode header,
    body.staff-mode #uploadForm,
    body.staff-mode .control-strip,
    body.staff-mode #configPanel,
    body.staff-mode #noticeBox,
    body.staff-mode .module-nav,
    body.staff-mode [data-module-view="pricing"],
    body.staff-mode [data-module-view="calendario"],
    body.staff-mode [data-module-view="biancheria"],
    body.staff-mode [data-module-view="rendiconto"] {
      display: none !important;
    }
    body.staff-mode main {
      padding-top: 16px;
    }
    body.staff-mode [data-module-view="pulizie"] {
      display: block;
    }
    .grid {
      display: grid;
      grid-template-columns: minmax(0, 1fr);
      gap: 18px;
    }
    .panel h2 {
      margin: 0;
      padding: 14px 16px;
      font-size: 18px;
      border-bottom: 1px solid var(--line);
    }
    .table-wrap { overflow: auto; }
    table { width: 100%; border-collapse: collapse; min-width: 760px; }
    th, td { padding: 10px 12px; border-bottom: 1px solid var(--line); text-align: left; white-space: nowrap; }
    th { font-size: 12px; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; background: #f9fbfc; }
    .num { text-align: right; font-variant-numeric: tabular-nums; }
    .badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      min-width: 76px;
      justify-content: center;
      padding: 4px 8px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 800;
    }
    .green { background: #e3f5ee; color: var(--green); }
    .yellow { background: #fff1cc; color: var(--yellow); }
    .red { background: #fde4df; color: var(--red); }
    tr.past-row {
      opacity: .48;
      background: #f6f8f9;
    }
    tr.warn-row {
      background: #fff9ea;
    }
    tr.critical-row {
      background: #fff1ef;
    }
    tr.past-row.warn-row, tr.past-row.critical-row {
      opacity: .56;
    }
    .empty {
      padding: 22px 16px;
      color: var(--muted);
    }
    .error {
      display: none;
      margin-bottom: 16px;
      padding: 12px 14px;
      border-radius: 8px;
      border: 1px solid #f1b8b1;
      background: #fff0ee;
      color: var(--red);
    }
    .notice {
      display: none;
      margin-bottom: 16px;
      padding: 12px 14px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: #fff;
      color: var(--ink);
      line-height: 1.45;
    }
    .notice strong {
      display: block;
      margin-bottom: 4px;
    }
    @media (max-width: 720px) {
      .kpis { grid-template-columns: 1fr; }
      .module-nav { grid-template-columns: 1fr; }
      .control-strip { grid-template-columns: 1fr; }
      .pricing-fields, .pricing-metrics { grid-template-columns: 1fr; }
      .calendar-head {
        display: grid;
      }
      .calendar-controls {
        justify-content: space-between;
      }
      .calendar-controls strong {
        min-width: 0;
      }
      .calendar-weekdays,
      .calendar-grid {
        min-width: 760px;
      }
      .calendar-panel {
        overflow-x: auto;
      }
      .upload, .settings { grid-template-columns: 1fr; }
      button { width: 100%; }
      .status-strip {
        align-items: start;
      }
      .status-strip.active { display: grid; }
      .cleaning-table {
        min-width: 0;
      }
      .cleaning-table thead {
        display: none;
      }
      .cleaning-table, .cleaning-table tbody, .cleaning-table tr, .cleaning-table td {
        display: block;
        width: 100%;
      }
      .cleaning-table tr {
        margin-bottom: 12px;
        border: 1px solid var(--line);
        border-radius: 8px;
        background: #fff;
        overflow: hidden;
      }
      .cleaning-table td {
        display: grid;
        grid-template-columns: 116px minmax(0, 1fr);
        gap: 10px;
        align-items: center;
        white-space: normal;
        border-bottom: 1px solid var(--line);
      }
      .cleaning-table td::before {
        content: attr(data-label);
        color: var(--muted);
        font-size: 11px;
        font-weight: 900;
        text-transform: uppercase;
      }
      .cleaning-table td:last-child {
        border-bottom: 0;
      }
      .metric-cell {
        text-align: left;
      }
      .date-chip {
        min-width: 0;
        width: 100%;
      }
      .legend {
        grid-template-columns: 1fr;
        font-size: 15px;
      }
    }
  </style>
</head>
<body>
  <header>
    <h1>Gestione appartamenti</h1>
    <p>Carica il file Excel una volta: vengono usati solo i fogli appartamento che iniziano con ID00. Poi scegli il modulo operativo.</p>
  </header>
  <main>
    <form class="upload" id="uploadForm">
      <input id="fileInput" name="file" type="file" accept=".xlsx" required>
      <div class="top-actions">
        <button id="uploadButton" type="submit">Carica e aggiorna</button>
      </div>
      <input type="hidden" name="apartments" id="apartmentsInput" value="[]">
      <input type="hidden" name="pricing" id="pricingInput" value="[]">
      <input type="hidden" name="deliveries" id="deliveriesInput" value="[]">
      <input type="hidden" name="staffUsers" id="staffUsersInput" value="[]">
    </form>
    <div class="control-strip">
      <div class="toolbar">
        <label class="field">Appartamento da consultare <select id="apartmentFilter"><option value="">Tutti gli appartamenti</option></select></label>
      </div>
      <button class="ghost" type="button" id="configToggle">Configura appartamenti</button>
      <button class="ghost" type="button" id="staffConfigToggle">Utenti staff</button>
    </div>
    <section class="panel config-panel" id="configPanel">
      <h2>Scheda configurazione appartamenti</h2>
      <div class="settings">
        <div class="subhead">Appartamenti, pricing base e biancheria partenza</div>
        <div class="apartment-panel table-wrap">
          <table>
            <thead>
              <tr>
                <th>ID foglio</th>
                <th>Nome visibile</th>
                <th>Indirizzo consegna</th>
                <th>Biancheria</th>
                <th>No dom.</th>
                <th class="num">Prezzo base</th>
                <th class="num">Prezzo weekend</th>
                <th class="num">Pulizie</th>
                <th class="num">Extra</th>
                <th class="num">Min</th>
                <th class="num">Max</th>
                <th class="num">Letto start</th>
                <th class="num">Bagno start</th>
                <th class="num">Tappeti start</th>
                <th>Prima fornitura</th>
                <th class="num">Soglia</th>
                <th class="num">Osp/letto</th>
                <th class="num">Osp/bagno</th>
                <th class="num">Tapp. std</th>
                <th class="num">Tapp. extra</th>
                <th>Col. ospiti</th>
                <th>Col. letto extra</th>
                <th></th>
              </tr>
            </thead>
            <tbody id="apartmentsRegistry"></tbody>
          </table>
        </div>
        <div class="delivery-actions">
          <button type="submit" form="uploadForm" id="refreshButton">Aggiorna calcoli</button>
        </div>
        <div class="subhead">Parametri default</div>
        <label class="field">Ospiti per kit letto <input name="bedGuests" form="uploadForm" type="number" min="1" step="1" value="2"></label>
        <label class="field">Ospiti per kit bagno <input name="bathGuests" form="uploadForm" type="number" min="1" step="1" value="1"></label>
        <label class="field">Tappetini standard <input name="matsStandard" form="uploadForm" type="number" min="0" step="1" value="1"></label>
        <label class="field">Tappetini extra <input name="matsExtra" form="uploadForm" type="number" min="0" step="1" value="0"></label>
        <div class="subhead">Consegne biancheria</div>
        <div class="delivery-panel table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data</th>
                <th>Appartamento</th>
                <th class="num">Kit letto</th>
                <th class="num">Kit bagno</th>
                <th class="num">Tappeti</th>
                <th>Note</th>
                <th></th>
              </tr>
            </thead>
            <tbody id="deliveries"></tbody>
          </table>
        </div>
        <input class="hidden-file" id="importDeliveriesFile" type="file" accept="application/json,.json,text/csv,.csv">
        <div class="delivery-actions">
          <button class="ghost" type="button" id="addDelivery">Aggiungi consegna</button>
          <button class="ghost" type="button" id="exportDeliveries">Scarica CSV consegne</button>
          <button class="ghost" type="button" id="importDeliveries">Importa archivio</button>
        </div>
        <div class="subhead" id="staffUsersSection">Utenti accesso</div>
        <div class="staff-panel table-wrap">
          <table>
            <thead>
              <tr>
                <th>Attivo</th>
                <th>Codice / nome</th>
                <th>Email username</th>
                <th>Telefono</th>
                <th>Ruolo</th>
                <th>Appartamenti visibili</th>
                <th>Moduli visibili</th>
                <th>Note</th>
                <th></th>
              </tr>
            </thead>
            <tbody id="staffUsers"></tbody>
          </table>
        </div>
        <div class="delivery-actions">
          <button class="ghost" type="button" id="addStaffUser">Aggiungi utente</button>
        </div>
      </div>
    </section>
    <div class="error" id="errorBox"></div>
    <div class="notice" id="noticeBox"></div>
    <section class="module-nav" aria-label="Moduli gestione">
      <button class="module-button active" type="button" data-module-target="pricing"><strong>Pricing</strong><span>Prezzi suggeriti giorno per giorno</span></button>
      <button class="module-button" type="button" data-module-target="calendario"><strong>Calendario</strong><span>Check-in e check-out da XLS</span></button>
      <button class="module-button" type="button" data-module-target="biancheria"><strong>Biancheria</strong><span>Scorte, consegne e consumi</span></button>
      <button class="module-button" type="button" data-module-target="pulizie"><strong>Pulizie</strong><span>Bozza planning operatori</span></button>
      <button class="module-button" type="button" data-module-target="rendiconto"><strong>Rendiconto</strong><span>Bozza entrate e uscite proprietari</span></button>
    </section>
    <section class="panel pricing-dashboard module-view active" data-module-view="pricing">
      <div class="pricing-head">
        <h2>Pricing giornaliero</h2>
        <span>Primo motore locale: occupazione, ADR, buchi e prezzo suggerito</span>
      </div>
      <div id="pricingCards" class="pricing-cards">
        <div class="empty">Carica un file Excel per generare prezzi e azioni giorno per giorno.</div>
      </div>
    </section>
    <section class="panel calendar-panel module-view" data-module-view="calendario">
      <div class="calendar-head">
        <h2>Calendario prenotazioni</h2>
        <div class="calendar-controls">
          <button class="ghost" type="button" id="calendarPrev" aria-label="Mese precedente">&lt;</button>
          <strong id="calendarMonthLabel">-</strong>
          <button class="ghost" type="button" id="calendarNext" aria-label="Mese successivo">&gt;</button>
        </div>
      </div>
      <div class="calendar-weekdays">
        <span>Lun</span><span>Mar</span><span>Mer</span><span>Gio</span><span>Ven</span><span>Sab</span><span>Dom</span>
      </div>
      <div class="calendar-grid" id="calendarGrid">
        <div class="empty">Carica un file Excel per vedere il calendario prenotazioni.</div>
      </div>
    </section>
    <section class="module-view" data-module-view="biancheria">
    <section class="kpis">
      <div class="kpi"><strong id="kpiApartments">0</strong><span>appartamenti letti</span></div>
      <div class="kpi"><strong id="kpiBookings">0</strong><span>prenotazioni importate</span></div>
      <div class="kpi"><strong id="kpiCritical">0</strong><span>righe con attenzione</span></div>
    </section>
    <section class="grid">
      <div class="panel">
        <h2>Dashboard scorte</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Appartamento</th>
                <th class="num">Partenza</th>
                <th class="num">Kit letto</th>
                <th class="num">Kit bagno</th>
                <th class="num">Tappeti</th>
                <th class="num">Da ordinare</th>
                <th>Consegnare entro</th>
                <th>Stato</th>
              </tr>
            </thead>
            <tbody id="apartmentsTable"><tr><td colspan="8" class="empty">Carica un file Excel per iniziare.</td></tr></tbody>
          </table>
        </div>
      </div>
      <div class="panel">
        <h2>Scadenziario consegne</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Appartamento</th>
                <th>#</th>
                <th>Ultima consegna</th>
                <th>Prossima ordinaria</th>
                <th>Consegna suggerita</th>
                <th class="num">Kit letto</th>
                <th class="num">Kit bagno</th>
                <th class="num">Tappeti</th>
                <th>Nota</th>
              </tr>
            </thead>
            <tbody id="scheduleTable"><tr><td colspan="9" class="empty">Carica un file Excel per generare lo scadenziario.</td></tr></tbody>
          </table>
        </div>
        <div class="inline-actions">
          <label class="field">Email fornitore <input id="supplierEmail" type="email" placeholder="fornitore@example.com"></label>
          <label class="field">WhatsApp fornitore <input id="supplierWhatsapp" type="tel" placeholder="393331234567"></label>
          <button class="ghost" type="button" id="copySupplierText">Copia testo ordine</button>
          <button class="ghost" type="button" id="whatsappSupplier">WhatsApp</button>
          <button type="button" id="emailSupplier">Prepara email</button>
        </div>
      </div>
      <div class="panel">
        <h2>Movimenti prenotazioni</h2>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Data in</th>
                <th>Appartamento</th>
                <th>Ospiti</th>
                <th>Letto extra</th>
                <th>Ospite</th>
                <th class="num">Uso letto</th>
                <th class="num">Uso bagno</th>
                <th class="num">Uso tappeti</th>
                <th class="num">Rim. letto</th>
                <th class="num">Rim. bagno</th>
                <th class="num">Rim. tappeti</th>
                <th>Stato</th>
              </tr>
            </thead>
            <tbody id="bookingsTable"><tr><td colspan="12" class="empty">Carica un file Excel per vedere tutti i consumi.</td></tr></tbody>
          </table>
        </div>
      </div>
    </section>
    </section>
    <section class="panel module-view" data-module-view="pulizie">
      <h2>Planning pulizie</h2>
      <div class="cleaning-tools">
        <div class="status-strip" id="cleaningUpdatedBox">
          <div>
            <span>Schema aggiornato al</span>
            <strong>-</strong>
          </div>
        </div>
        <div class="legend">
          <span>1 Letto = 1 kit letto</span>
          <span>1 Ospite = 1 kit bagno</span>
          <span>1 kit bagno = asciugamano grande + medio + piccolo</span>
        </div>
      </div>
      <div class="table-wrap">
        <table class="cleaning-table">
          <thead>
            <tr>
              <th><span class="th-chip cleaning">Quando<br>pulire</span></th>
              <th>Priorita</th>
              <th>Appartamento</th>
              <th>Ospite uscente</th>
              <th><span class="th-chip checkout">Uscita<br>ospiti</span></th>
              <th class="metric-cell"><span class="th-chip metric">Ospiti<br>da preparare</span></th>
              <th class="metric-cell"><span class="th-chip metric">Letti<br>da fare</span></th>
              <th><span class="th-chip arrival">Arrivo<br>successivo</span></th>
              <th>Canale</th>
              <th>Note</th>
            </tr>
          </thead>
          <tbody id="cleaningTable"><tr><td colspan="10" class="empty">Carica un file Excel per generare il planning pulizie.</td></tr></tbody>
        </table>
      </div>
    </section>
    <section class="panel module-view coming-soon" data-module-view="rendiconto">
      <h2>Rendiconto proprietari</h2>
      <p>Qui costruiremo il rendiconto per proprietari con entrate, uscite, pulizie, extra e movimenti filtrati per appartamento.</p>
    </section>
  </main>
  <script>
    const form = document.getElementById('uploadForm');
    const button = document.getElementById('uploadButton');
    const errorBox = document.getElementById('errorBox');
    const noticeBox = document.getElementById('noticeBox');
    const deliveriesBox = document.getElementById('deliveries');
    const deliveriesInput = document.getElementById('deliveriesInput');
    const staffUsersBox = document.getElementById('staffUsers');
    const staffUsersInput = document.getElementById('staffUsersInput');
    const apartmentsRegistry = document.getElementById('apartmentsRegistry');
    const apartmentsInput = document.getElementById('apartmentsInput');
    const pricingInput = document.getElementById('pricingInput');
    const pricingCards = document.getElementById('pricingCards');
    const addApartment = document.getElementById('addApartment');
    const apartmentFilter = document.getElementById('apartmentFilter');
    const configToggle = document.getElementById('configToggle');
    const staffConfigToggle = document.getElementById('staffConfigToggle');
    const staffUsersSection = document.getElementById('staffUsersSection');
    const configPanel = document.getElementById('configPanel');
    const addDelivery = document.getElementById('addDelivery');
    const refreshButton = document.getElementById('refreshButton');
    const exportDeliveries = document.getElementById('exportDeliveries');
    const importDeliveries = document.getElementById('importDeliveries');
    const importDeliveriesFile = document.getElementById('importDeliveriesFile');
    const addStaffUser = document.getElementById('addStaffUser');
    const supplierEmail = document.getElementById('supplierEmail');
    const supplierWhatsapp = document.getElementById('supplierWhatsapp');
    const copySupplierText = document.getElementById('copySupplierText');
    const whatsappSupplier = document.getElementById('whatsappSupplier');
    const emailSupplier = document.getElementById('emailSupplier');
    const moduleButtons = [...document.querySelectorAll('[data-module-target]')];
    const moduleViews = [...document.querySelectorAll('[data-module-view]')];
    const cleaningUpdatedBox = document.getElementById('cleaningUpdatedBox');
    const calendarGrid = document.getElementById('calendarGrid');
    const calendarMonthLabel = document.getElementById('calendarMonthLabel');
    const calendarPrev = document.getElementById('calendarPrev');
    const calendarNext = document.getElementById('calendarNext');
    const fmt = new Intl.DateTimeFormat('it-IT');
    const fmtDateTime = new Intl.DateTimeFormat('it-IT', { dateStyle: 'short', timeStyle: 'short' });
    const fmtMonth = new Intl.DateTimeFormat('it-IT', { month: 'long', year: 'numeric' });
    const archiveKey = 'biancheria-deliveries-v1';
    const apartmentsKey = 'biancheria-apartments-v1';
    const pricingKey = 'pricing-apartments-v1';
    const staffUsersKey = 'pulizie-staff-users-v1';
    const params = new URLSearchParams(window.location.search);
    const staffMode = params.get('staff') === '1' || params.get('view') === 'pulizie';
    let lastData = null;
    let lastSchedule = [];
    let calendarMonth = new Date();
    calendarMonth = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth(), 1);

    function esc(value) {
      return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    }

    function fmtDate(value) {
      if (!value) return '';
      const d = new Date(value + 'T00:00:00');
      return Number.isNaN(d.getTime()) ? esc(value) : fmt.format(d);
    }

    function fmtDateTimeValue(value) {
      if (!value) return '-';
      const d = new Date(value);
      return Number.isNaN(d.getTime()) ? esc(value) : fmtDateTime.format(d);
    }

    function csvEscape(value) {
      const text = String(value ?? '');
      return /[;"\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
    }

    function deliveriesToCsv(rows) {
      const header = ['Data', 'Appartamento', 'Kit letto', 'Kit bagno', 'Tappeti', 'Note'];
      const lines = rows.map(row => [
        row.date, row.apartment, row.bed, row.bath, row.mats, row.note
      ].map(csvEscape).join(';'));
      return [header.join(';'), ...lines].join('\n');
    }

    function csvToDeliveries(text) {
      const lines = text.split(/\r?\n/).filter(Boolean);
      return lines.slice(1).map(line => {
        const parts = [];
        let current = '';
        let quoted = false;
        for (let i = 0; i < line.length; i++) {
          const ch = line[i];
          if (ch === '"' && line[i + 1] === '"') {
            current += '"';
            i += 1;
          } else if (ch === '"') {
            quoted = !quoted;
          } else if (ch === ';' && !quoted) {
            parts.push(current);
            current = '';
          } else {
            current += ch;
          }
        }
        parts.push(current);
        return {
          date: parts[0] || '',
          apartment: parts[1] || '',
          bed: parts[2] || '0',
          bath: parts[3] || '0',
          mats: parts[4] || '0',
          note: parts[5] || '',
        };
      });
    }

    function badge(status) {
      const label = status === 'red' ? 'Critico' : status === 'yellow' ? 'Attenzione' : 'Ok';
      return `<span class="badge ${status}">${label}</span>`;
    }

    function stockCell(value, threshold) {
      const cls = value < 0 ? 'red' : value <= threshold ? 'yellow' : 'green';
      return `<span class="badge ${cls}">${esc(value)}</span>`;
    }

    function setActiveModule(moduleName) {
      moduleButtons.forEach(button => {
        button.classList.toggle('active', button.dataset.moduleTarget === moduleName);
      });
      moduleViews.forEach(view => {
        view.classList.toggle('active', view.dataset.moduleView === moduleName);
      });
    }

    function isoDate(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    function channelCode(value) {
      const text = String(value || '').toLowerCase();
      if (text.includes('airbnb') || text === 'ab') return 'Ab';
      if (text.includes('booking') || text === 'bk') return 'Bk';
      return 'Dr';
    }

    function renderCalendar(data) {
      if (!calendarGrid || !calendarMonthLabel) return;
      const selectedApartment = apartmentFilter.value || '';
      const rows = (data.reservations || []).filter(row => !selectedApartment || row.apartment === selectedApartment);
      calendarMonthLabel.textContent = fmtMonth.format(calendarMonth);
      const first = new Date(calendarMonth);
      const weekdayOffset = (first.getDay() + 6) % 7;
      const start = new Date(first);
      start.setDate(first.getDate() - weekdayOffset);
      const end = new Date(start);
      end.setDate(start.getDate() + 42);
      const cells = [];
      for (let index = 0; index < 42; index++) {
        const current = new Date(start);
        current.setDate(start.getDate() + index);
        cells.push(`
          <div class="calendar-day ${current.getMonth() === calendarMonth.getMonth() ? '' : 'outside'}" style="grid-row:${Math.floor(index / 7) + 1}">
            <span class="calendar-day-number">${current.getDate()}</span>
          </div>
        `);
      }
      const weeklySegments = Array.from({ length: 6 }, () => []);
      rows.forEach(row => {
        if (!row.checkIn || !row.checkOut) return;
        const bookingStart = new Date(row.checkIn + 'T00:00:00');
        const bookingEnd = new Date(row.checkOut + 'T00:00:00');
        if (bookingEnd <= start || bookingStart >= end) return;
        for (let week = 0; week < 6; week++) {
          const weekStart = new Date(start);
          weekStart.setDate(start.getDate() + week * 7);
          const weekEnd = new Date(weekStart);
          weekEnd.setDate(weekStart.getDate() + 7);
          const segmentStart = bookingStart > weekStart ? bookingStart : weekStart;
          const segmentEnd = bookingEnd < weekEnd ? bookingEnd : weekEnd;
          if (segmentStart >= segmentEnd) continue;
          weeklySegments[week].push({
            startCol: Math.floor((segmentStart - weekStart) / 86400000) + 1,
            span: Math.floor((segmentEnd - segmentStart) / 86400000),
            continuesLeft: bookingStart < weekStart,
            continuesRight: bookingEnd > weekEnd,
            label: `${channelCode(row.ota)} ${row.guestName || 'Ospite'}`,
            apartment: row.apartmentLabel || row.apartment,
          });
        }
      });
      const bars = [];
      weeklySegments.forEach((segments, week) => {
        const laneEnds = [];
        segments
          .sort((a, b) => a.startCol - b.startCol || b.span - a.span || a.apartment.localeCompare(b.apartment, 'it'))
          .forEach(segment => {
            let lane = laneEnds.findIndex(endCol => segment.startCol > endCol);
            if (lane === -1) {
              lane = laneEnds.length;
              laneEnds.push(0);
            }
            laneEnds[lane] = segment.startCol + segment.span - 1;
            const classes = [
              'calendar-booking',
              segment.continuesLeft ? 'continues-left' : '',
              segment.continuesRight ? 'continues-right' : '',
            ].filter(Boolean).join(' ');
            bars.push(`
              <div class="${classes}" style="grid-column:${segment.startCol} / span ${segment.span}; grid-row:${week + 1}; margin-top:${29 + lane * 32}px">
                ${esc(segment.label)} <small>${esc(segment.apartment)}</small>
              </div>
            `);
          });
      });
      calendarGrid.innerHTML = cells.join('') + bars.join('');
    }

    function defaultApartmentValues() {
      const valueOf = (name, fallback) => {
        const input = document.querySelector(`[name="${name}"]`);
        return input ? input.value || fallback : fallback;
      };
      return {
        bedGuests: valueOf('bedGuests', '2'),
        bathGuests: valueOf('bathGuests', '1'),
        matsStandard: valueOf('matsStandard', '1'),
        matsExtra: valueOf('matsExtra', '0'),
        startBed: '0',
        startBath: '0',
        startMats: '0',
        startDate: '',
        threshold: '2',
        useLaundry: '',
        avoidSundayCleaning: '',
        basePrice: '90',
        weekendPrice: '110',
        cleaningFee: '0',
        extraFee: '0',
        minPrice: '65',
        maxPrice: '180',
        guestCol: '',
        extraBedCol: '',
      };
    }

    function addDeliveryRow(values = {}) {
      const row = document.createElement('tr');
      row.className = 'delivery-row';
      row.innerHTML = `
        <td><input data-delivery="date" type="date"></td>
        <td><input data-delivery="apartment" type="text" placeholder="vuoto = tutti"></td>
        <td><input data-delivery="bed" type="number" min="0" step="1" value="0"></td>
        <td><input data-delivery="bath" type="number" min="0" step="1" value="0"></td>
        <td><input data-delivery="mats" type="number" min="0" step="1" value="0"></td>
        <td><input data-delivery="note" type="text"></td>
        <td><button class="ghost tiny" type="button">Rimuovi</button></td>
      `;
      row.querySelectorAll('[data-delivery]').forEach(input => {
        input.value = values[input.dataset.delivery] ?? input.value;
        input.addEventListener('input', saveDeliveriesArchive);
      });
      row.querySelector('button').addEventListener('click', () => {
        row.remove();
        saveDeliveriesArchive();
      });
      deliveriesBox.appendChild(row);
    }

    function addStaffUserRow(values = {}) {
      const row = document.createElement('tr');
      row.className = 'staff-user-row';
      row.innerHTML = `
        <td><input data-staff="active" type="checkbox" value="1"></td>
        <td><input data-staff="code" type="text" placeholder="Maria / STAFF01"></td>
        <td><input data-staff="email" type="email" placeholder="maria@example.com"></td>
        <td><input data-staff="phone" type="tel" placeholder="+393331234567"></td>
        <td>
          <select data-staff="role">
            <option value="pulizie">Pulizie</option>
            <option value="proprietario">Proprietario</option>
            <option value="admin">Admin</option>
          </select>
        </td>
        <td><input data-staff="apartments" type="text" placeholder="vuoto = tutti, oppure ID001, ID002"></td>
        <td><input data-staff="modules" type="text" placeholder="pulizie, calendario"></td>
        <td><input data-staff="note" type="text" placeholder="note interne"></td>
        <td><button class="ghost tiny" type="button">Rimuovi</button></td>
      `;
      row.querySelectorAll('[data-staff]').forEach(input => {
        if (input.type === 'checkbox') {
          input.checked = !('active' in values) || ['1', 'true', 'on', 'si', 'sì', true].includes(values.active);
        } else {
          input.value = values[input.dataset.staff] ?? (input.dataset.staff === 'role' ? 'pulizie' : '');
        }
        input.addEventListener('input', saveStaffUsersArchive);
        input.addEventListener('change', saveStaffUsersArchive);
      });
      row.querySelector('button').addEventListener('click', () => {
        row.remove();
        saveStaffUsersArchive();
      });
      staffUsersBox.appendChild(row);
    }

    function getStaffUsers() {
      return [...staffUsersBox.querySelectorAll('.staff-user-row')].map(row => {
        const data = {};
        row.querySelectorAll('[data-staff]').forEach(input => data[input.dataset.staff] = input.type === 'checkbox' ? (input.checked ? '1' : '') : input.value);
        return data;
      }).filter(row => row.code || row.email || row.phone);
    }

    function serializeStaffUsers() {
      staffUsersInput.value = JSON.stringify(getStaffUsers());
    }

    function saveStaffUsersArchive() {
      const rows = getStaffUsers();
      localStorage.setItem(staffUsersKey, JSON.stringify(rows));
      serializeStaffUsers();
    }

    function restoreStaffUsersArchive() {
      try {
        const rows = JSON.parse(localStorage.getItem(staffUsersKey) || '[]');
        staffUsersBox.innerHTML = '';
        (Array.isArray(rows) ? rows : []).forEach(row => addStaffUserRow(row));
      } catch {
        staffUsersBox.innerHTML = '';
      }
      serializeStaffUsers();
    }

    function addApartmentRow(values = {}) {
      const defaults = defaultApartmentValues();
      values = { ...defaults, ...values };
      const row = document.createElement('tr');
      row.className = 'apartment-row';
      row.innerHTML = `
        <td><input data-apartment="id" type="text" placeholder="ID006_Via Fondazza" readonly></td>
        <td><input data-apartment="label" type="text" placeholder="Via Fondazza"></td>
        <td><input data-apartment="address" type="text" placeholder="Via..., Bologna"></td>
        <td><input data-apartment="useLaundry" type="checkbox" value="1" title="Includi nel modulo biancheria"></td>
        <td><input data-apartment="avoidSundayCleaning" type="checkbox" value="1" title="Evita pulizie di domenica se possibile"></td>
        <td><input data-apartment="basePrice" type="number" min="0" step="1" value="90"></td>
        <td><input data-apartment="weekendPrice" type="number" min="0" step="1" value="110"></td>
        <td><input data-apartment="cleaningFee" type="number" min="0" step="1" value="0"></td>
        <td><input data-apartment="extraFee" type="number" min="0" step="1" value="0"></td>
        <td><input data-apartment="minPrice" type="number" min="0" step="1" value="65"></td>
        <td><input data-apartment="maxPrice" type="number" min="0" step="1" value="180"></td>
        <td><input data-apartment="startBed" type="number" min="0" step="1" value="0"></td>
        <td><input data-apartment="startBath" type="number" min="0" step="1" value="0"></td>
        <td><input data-apartment="startMats" type="number" min="0" step="1" value="0"></td>
        <td><input data-apartment="startDate" type="date" title="Da questa data in poi vengono considerati solo i check-in validi"></td>
        <td><input data-apartment="threshold" type="number" min="0" step="1" value="2"></td>
        <td><input data-apartment="bedGuests" type="number" min="1" step="1" value="2"></td>
        <td><input data-apartment="bathGuests" type="number" min="1" step="1" value="1"></td>
        <td><input data-apartment="matsStandard" type="number" min="0" step="1" value="1"></td>
        <td><input data-apartment="matsExtra" type="number" min="0" step="1" value="0"></td>
        <td><input data-apartment="guestCol" type="text" placeholder="es. H"></td>
        <td><input data-apartment="extraBedCol" type="text" placeholder="es. AN"></td>
        <td><button class="ghost tiny" type="button">Rimuovi</button></td>
      `;
      row.querySelectorAll('[data-apartment]').forEach(input => {
        if (input.type === 'checkbox') {
          input.checked = ['1', 'true', 'on', 'si', 'sì', true].includes(values[input.dataset.apartment]);
        } else {
          input.value = values[input.dataset.apartment] ?? input.value;
        }
        input.addEventListener('input', saveApartmentsArchive);
        input.addEventListener('change', saveApartmentsArchive);
      });
      row.querySelector('button').addEventListener('click', () => {
        row.remove();
        saveApartmentsArchive();
      });
      apartmentsRegistry.appendChild(row);
    }

    function sheetNameToLabel(name) {
      return String(name || '')
        .replace(/^(?:apt|app|id)?\\s*[\\-_ ]*\\d+[A-Za-z]?\\s*[\\-_ ]+/i, '')
        .replace(/[_-]+/g, ' ')
        .replace(/\\s+/g, ' ')
        .trim() || String(name || '');
    }

    function ensureApartmentRowsFromWorkbook(apartments = []) {
      const saved = new Map(getStoredApartments().map(row => [row.id, row]));
      apartmentsRegistry.innerHTML = '';
      if (!apartments.length) {
        apartmentsRegistry.innerHTML = '<tr><td colspan="23" class="empty">Nessun foglio appartamento trovato nel file caricato.</td></tr>';
        apartmentsInput.value = JSON.stringify(getStoredApartments());
        return;
      }
      apartments.forEach(apartment => {
        if (!apartment.name) return;
        const stored = saved.get(apartment.name) || {};
        addApartmentRow({
          ...stored,
          id: apartment.name,
          label: stored.label || apartment.label || sheetNameToLabel(apartment.name),
          address: stored.address || apartment.address || '',
          bedGuests: stored.bedGuests || apartment.settings?.bedGuests || '2',
          bathGuests: stored.bathGuests || apartment.settings?.bathGuests || '1',
          matsStandard: stored.matsStandard || apartment.settings?.matsStandard || '1',
          matsExtra: stored.matsExtra || apartment.settings?.matsExtra || '0',
          startBed: stored.startBed || apartment.settings?.startBed || '0',
          startBath: stored.startBath || apartment.settings?.startBath || '0',
          startMats: stored.startMats || apartment.settings?.startMats || '0',
          startDate: stored.startDate || apartment.settings?.startDate || '',
          threshold: stored.threshold || apartment.settings?.threshold || '2',
          useLaundry: stored.useLaundry || apartment.settings?.useLaundry || '',
          avoidSundayCleaning: stored.avoidSundayCleaning || apartment.settings?.avoidSundayCleaning || '',
          basePrice: stored.basePrice || '90',
          weekendPrice: stored.weekendPrice || '110',
          cleaningFee: stored.cleaningFee || '0',
          extraFee: stored.extraFee || '0',
          minPrice: stored.minPrice || '65',
          maxPrice: stored.maxPrice || '180',
          guestCol: stored.guestCol || apartment.settings?.guestCol || '',
          extraBedCol: stored.extraBedCol || apartment.settings?.extraBedCol || '',
        });
      });
      saveApartmentsArchive();
      serializeApartments();
    }

    function selectedApartmentRow() {
      const selected = apartmentFilter.value;
      if (!selected) return null;
      return [...apartmentsRegistry.querySelectorAll('.apartment-row')].find(row => {
        const id = row.querySelector('[data-apartment="id"]');
        return id && id.value === selected;
      });
    }

    function syncVisibleSettingsToSelectedApartment() {
      const row = selectedApartmentRow();
      if (!row) return;
      ['bedGuests', 'bathGuests', 'matsStandard', 'matsExtra', 'startBed', 'startBath', 'startMats', 'startDate', 'threshold'].forEach(key => {
        const source = document.querySelector(`[name="${key}"]`);
        const target = row.querySelector(`[data-apartment="${key}"]`);
        if (source && target) target.value = source.value;
      });
      saveApartmentsArchive();
    }

    function getApartmentsRegistry() {
      return [...apartmentsRegistry.querySelectorAll('.apartment-row')].map(row => {
        const data = {};
        row.querySelectorAll('[data-apartment]').forEach(input => data[input.dataset.apartment] = input.type === 'checkbox' ? (input.checked ? '1' : '') : input.value);
        return data;
      }).filter(row => row.id || row.label || row.address);
    }

    function serializeApartments() {
      const visibleRows = getApartmentsRegistry();
      apartmentsInput.value = JSON.stringify(visibleRows.length ? visibleRows : getStoredApartments());
    }

    function saveApartmentsArchive() {
      const rows = getApartmentsRegistry();
      if (rows.length) localStorage.setItem(apartmentsKey, JSON.stringify(rows));
    }

    function getStoredApartments() {
      try {
        const rows = JSON.parse(localStorage.getItem(apartmentsKey) || '[]');
        return Array.isArray(rows) ? rows : [];
      } catch {
        return [];
      }
    }

    function getStoredPricing() {
      try {
        const rows = JSON.parse(localStorage.getItem(pricingKey) || '[]');
        return Array.isArray(rows) ? rows : [];
      } catch {
        return [];
      }
    }

    function pricingDefaults(item = {}) {
      const base = Number(item?.settings?.basePrice ?? 90);
      return {
        id: item.apartment || item.id || '',
        basePrice: String(Math.round(base)),
        weekendPrice: String(Math.round(item?.settings?.weekendPrice ?? base * 1.18)),
        cleaningFee: String(Math.round(item?.settings?.cleaningFee ?? 0)),
        extraFee: String(Math.round(item?.settings?.extraFee ?? 0)),
        minPrice: String(Math.round(item?.settings?.minPrice ?? Math.max(35, base * 0.72))),
        maxPrice: String(Math.round(item?.settings?.maxPrice ?? base * 1.75)),
        notes: item?.settings?.notes || '',
      };
    }

    function collectPricingSettings() {
      const configRows = getApartmentsRegistry().filter(row => row.id).map(row => ({
        id: row.id,
        basePrice: row.basePrice,
        weekendPrice: row.weekendPrice,
        cleaningFee: row.cleaningFee,
        extraFee: row.extraFee,
        minPrice: row.minPrice,
        maxPrice: row.maxPrice,
        notes: row.notes || '',
      }));
      if (configRows.length) return configRows;
      const rows = [...pricingCards.querySelectorAll('.pricing-card')].map(card => {
        const data = { id: card.dataset.apartment || '' };
        card.querySelectorAll('[data-pricing]').forEach(input => data[input.dataset.pricing] = input.value);
        return data;
      }).filter(row => row.id);
      if (rows.length) return rows;
      return getStoredPricing();
    }

    function serializePricing() {
      pricingInput.value = JSON.stringify(collectPricingSettings());
    }

    function syncPricingCardsToConfigRows() {
      pricingCards.querySelectorAll('.pricing-card').forEach(card => {
        const id = card.dataset.apartment || '';
        if (!id) return;
        const row = [...apartmentsRegistry.querySelectorAll('.apartment-row')].find(item => {
          const input = item.querySelector('[data-apartment="id"]');
          return input && input.value === id;
        });
        if (!row) return;
        card.querySelectorAll('[data-pricing]').forEach(input => {
          const target = row.querySelector(`[data-apartment="${input.dataset.pricing}"]`);
          if (target) target.value = input.value;
        });
      });
    }

    function savePricingArchive() {
      syncPricingCardsToConfigRows();
      saveApartmentsArchive();
      const current = collectPricingSettings();
      const merged = new Map(getStoredPricing().map(row => [row.id, row]));
      current.forEach(row => merged.set(row.id, row));
      const rows = [...merged.values()];
      localStorage.setItem(pricingKey, JSON.stringify(rows));
      pricingInput.value = JSON.stringify(rows);
    }

    function restoreApartmentsArchive() {
      apartmentsRegistry.innerHTML = `<tr><td colspan="23" class="empty">Carica l'XLS: gli appartamenti verranno letti automaticamente dai nomi dei fogli.</td></tr>`;
      apartmentsInput.value = JSON.stringify(getStoredApartments());
    }

    function getDeliveries() {
      const rows = [...deliveriesBox.querySelectorAll('.delivery-row')].map(row => {
        const data = {};
        row.querySelectorAll('[data-delivery]').forEach(input => data[input.dataset.delivery] = input.value);
        return data;
      }).filter(row => row.date || Number(row.bed) || Number(row.bath) || Number(row.mats));
      return rows;
    }

    function serializeDeliveries() {
      deliveriesInput.value = JSON.stringify(getDeliveries());
    }

    function saveDeliveriesArchive() {
      localStorage.setItem(archiveKey, JSON.stringify(getDeliveries()));
    }

    function loadDeliveriesArchive(rows) {
      deliveriesBox.innerHTML = '';
      rows.forEach(row => addDeliveryRow(row));
      saveDeliveriesArchive();
    }

    function restoreDeliveriesArchive() {
      try {
        const rows = JSON.parse(localStorage.getItem(archiveKey) || '[]');
        loadDeliveriesArchive(Array.isArray(rows) ? rows : []);
      } catch {
        loadDeliveriesArchive([]);
      }
    }

    function renderPricing(data) {
      const selectedApartment = apartmentFilter.value || '';
      const stored = collectPricingSettings();
      const storedById = new Map(stored.map(row => [row.id, row]));
      const rows = (data.pricing || []).filter(item => !selectedApartment || item.apartment === selectedApartment);
      if (!rows.length) {
        pricingCards.innerHTML = '<div class="empty">Nessun dato pricing disponibile per la selezione corrente.</div>';
        serializePricing();
        return;
      }

      pricingCards.innerHTML = rows.map(item => {
        const saved = storedById.get(item.apartment) || {};
        const defaults = pricingDefaults(item);
        const values = { ...defaults, ...saved, id: item.apartment };
        const calendarRows = (item.calendar || []).map(day => {
          const finalCell = day.occupied
            ? `<span class="occupied-price">${esc(day.guestName || 'Occupato')}</span>`
            : `<span class="price-pill">€ ${esc(day.finalPrice)}</span>`;
          const current = day.currentPrice ? `€ ${esc(day.currentPrice)}` : '';
          return `
            <tr class="${day.occupied ? 'past-row' : day.action === 'spingere' ? 'warn-row' : ''}">
              <td>${fmtDate(day.date)}</td>
              <td>${esc(day.weekday)}</td>
              <td class="num">${current}</td>
              <td class="num">€ ${esc(day.basePrice)}</td>
              <td class="num">${finalCell}</td>
              <td>${esc(day.action)}</td>
              <td>${esc(day.reason)}</td>
            </tr>
          `;
        }).join('');
        return `
          <article class="pricing-card" data-apartment="${esc(item.apartment)}">
            <h3>${esc(item.apartmentLabel || item.apartment)}</h3>
            <div class="pricing-metrics">
              <div><strong>${item.adr ? `€ ${esc(item.adr)}` : '-'}</strong><span>ADR da XLS</span></div>
              <div><strong>${esc(item.occupancy30)}%</strong><span>occupazione 30 gg</span></div>
              <div><strong>${item.averageSuggested ? `€ ${esc(item.averageSuggested)}` : '-'}</strong><span>media suggerita</span></div>
            </div>
            <div class="pricing-fields">
              <label class="field">Prezzo base <input data-pricing="basePrice" type="number" min="0" step="1" value="${esc(values.basePrice)}"></label>
              <label class="field">Prezzo weekend <input data-pricing="weekendPrice" type="number" min="0" step="1" value="${esc(values.weekendPrice)}"></label>
              <label class="field">Pulizie <input data-pricing="cleaningFee" type="number" min="0" step="1" value="${esc(values.cleaningFee)}"></label>
              <label class="field">Extra <input data-pricing="extraFee" type="number" min="0" step="1" value="${esc(values.extraFee)}"></label>
              <label class="field">Minimo <input data-pricing="minPrice" type="number" min="0" step="1" value="${esc(values.minPrice)}"></label>
              <label class="field">Massimo <input data-pricing="maxPrice" type="number" min="0" step="1" value="${esc(values.maxPrice)}"></label>
              <label class="field">Note pricing <input data-pricing="notes" type="text" value="${esc(values.notes)}"></label>
            </div>
            <div class="pricing-table">
              <table>
                <thead>
                  <tr>
                    <th>Data</th>
                    <th>Giorno</th>
                    <th class="num">Attuale</th>
                    <th class="num">Base</th>
                    <th class="num">Da impostare</th>
                    <th>Azione</th>
                    <th>Motivo</th>
                  </tr>
                </thead>
                <tbody>${calendarRows}</tbody>
              </table>
            </div>
          </article>
        `;
      }).join('');

      pricingCards.querySelectorAll('[data-pricing]').forEach(input => {
        input.addEventListener('input', savePricingArchive);
      });
      savePricingArchive();
    }

    function renderCleaning(data) {
      const selectedApartment = apartmentFilter.value || '';
      const rows = (data.cleaningSchedule || []).filter(row => !selectedApartment || row.apartment === selectedApartment);
      const html = rows.map(row => {
        const priorityClass = row.priority === 'Alta' ? 'critical-row' : row.priority === 'Media' ? 'warn-row' : '';
        return `
          <tr class="${priorityClass}">
            <td data-label="Quando pulire"><span class="date-chip cleaning">${fmtDate(row.date)}<span>intervento</span></span></td>
            <td data-label="Priorita">${esc(row.priority || '')}</td>
            <td data-label="Appartamento">${esc(row.apartmentLabel || row.apartment)}</td>
            <td data-label="Ospite uscente">${esc(row.guestName || '')}</td>
            <td data-label="Uscita ospiti"><span class="date-chip checkout">${fmtDate(row.checkout)}<span>uscita <b class="time-pill">${esc(row.checkoutTime || '10:30')}</b></span></span></td>
            <td data-label="Ospiti da preparare" class="metric-cell">${row.nextGuests ? `<span class="metric-chip"><strong>${esc(row.nextGuests)}</strong><span>ospiti</span></span>` : ''}</td>
            <td data-label="Letti da fare" class="metric-cell">${row.bedsToPrepare ? `<span class="metric-chip"><strong>${esc(row.bedsToPrepare)}</strong><span>letti</span></span>` : ''}</td>
            <td data-label="Arrivo successivo">${row.nextCheckin ? `<span class="date-chip arrival">${fmtDate(row.nextCheckin)}<span>arrivo <b class="time-pill">${esc(row.nextCheckinTime || '16:00')}</b></span></span>` : ''}</td>
            <td data-label="Canale">${esc(row.ota || '')}</td>
            <td data-label="Note">${esc(row.note || '')}</td>
          </tr>
        `;
      }).join('');
      document.getElementById('cleaningTable').innerHTML = html || '<tr><td colspan="10" class="empty">Nessuna pulizia trovata per la selezione corrente.</td></tr>';
    }

    function render(data) {
      lastData = data;
      if (cleaningUpdatedBox) {
        cleaningUpdatedBox.classList.add('active');
        const strong = cleaningUpdatedBox.querySelector('strong');
        if (strong) strong.textContent = fmtDateTimeValue(data.generatedAt);
      }
      const apartmentIndex = new Map((data.apartments || []).map(item => [item.name, item]));
      const workbookApartments = (data.workbookSheets || data.apartments || []).map(item => ({
        ...(apartmentIndex.get(item.name) || {}),
        name: item.name,
        label: (apartmentIndex.get(item.name) || item).label || sheetNameToLabel(item.name),
        address: (apartmentIndex.get(item.name) || item).address || '',
        settings: (apartmentIndex.get(item.name) || item).settings || {},
        parseStatus: item.parseStatus || data.parseStatus?.[item.name] || { ok: true, message: '' },
      }));
      noticeBox.style.display = 'block';
      if (workbookApartments.length) {
        const names = workbookApartments.map(item => item.name).join(', ');
        const ignored = (data.ignoredSheets || []).length ? `<br>Ignorati: ${esc((data.ignoredSheets || []).join(', '))}` : '';
        const warnings = workbookApartments
          .filter(item => item.parseStatus && !item.parseStatus.ok)
          .map(item => `${item.label || item.name}: ${item.parseStatus.message}`);
        const warningText = warnings.length ? `<br>Da controllare: ${esc(warnings.join(' | '))}` : '';
        noticeBox.innerHTML = `<strong>File letto: ${workbookApartments.length} appartamenti trovati.</strong>${esc(names)}${ignored}${warningText}`;
      } else {
        const ignored = (data.ignoredSheets || []).length ? ` Fogli ignorati: ${esc((data.ignoredSheets || []).join(', '))}.` : '';
        noticeBox.innerHTML = `<strong>File letto, ma nessun foglio appartamento ID00 trovato.</strong>Vengono usati solo i fogli il cui nome inizia con ID00.${ignored}`;
      }
      ensureApartmentRowsFromWorkbook(workbookApartments);
      const threshold = Number(data.summary.threshold ?? 2);
      const selectedApartment = apartmentFilter.value || '';
      const visibleApartments = (data.apartments || []).filter(a => !selectedApartment || a.name === selectedApartment);
      lastSchedule = (data.deliverySchedule || []).filter(row => !selectedApartment || row.apartment === selectedApartment);
      const visibleBookings = (data.bookings || []).filter(row => !selectedApartment || row.apartment === selectedApartment);
      apartmentFilter.innerHTML = '<option value="">Tutti gli appartamenti</option>' + workbookApartments.map(a => `<option value="${esc(a.name)}">${esc(a.label || a.name)}</option>`).join('');
      apartmentFilter.value = selectedApartment;
      document.getElementById('kpiApartments').textContent = data.summary.apartments;
      document.getElementById('kpiBookings').textContent = visibleBookings.filter(row => row.type === 'booking').length;
      document.getElementById('kpiCritical').textContent = visibleBookings.filter(row => row.type === 'booking' && row.status !== 'green').length;
      const selectedMeta = selectedApartment ? (data.apartments || []).find(a => a.name === selectedApartment) : null;
      if (selectedMeta) {
        [
          ['bedGuests', selectedMeta.settings.bedGuests],
          ['bathGuests', selectedMeta.settings.bathGuests],
          ['matsStandard', selectedMeta.settings.matsStandard],
          ['matsExtra', selectedMeta.settings.matsExtra],
        ].forEach(([name, value]) => {
          const input = document.querySelector(`[name="${name}"]`);
          if (input) input.value = value;
        });
      }

      const apartmentRows = visibleApartments.map(a => `
        <tr>
          <td>${esc(a.label || a.name)}<br><span style="color: var(--muted); font-size: 12px;">${esc(a.name)}</span></td>
          <td class="num">${esc(a.initial.bed)} / ${esc(a.initial.bath)} / ${esc(a.initial.mats)}</td>
          <td class="num">${stockCell(a.stock.bed, a.settings.threshold)}</td>
          <td class="num">${stockCell(a.stock.bath, a.settings.threshold)}</td>
          <td class="num">${stockCell(a.stock.mats, a.settings.threshold)}</td>
          <td class="num">${esc(a.need.bed)} / ${esc(a.need.bath)} / ${esc(a.need.mats)}</td>
          <td>${fmtDate(a.deliverByDate)}</td>
          <td>${badge(a.status)}</td>
        </tr>
      `).join('');
      document.getElementById('apartmentsTable').innerHTML = apartmentRows || '<tr><td colspan="8" class="empty">Nessun foglio appartamento riconosciuto.</td></tr>';

      const scheduleRows = lastSchedule.map(row => {
        const hasOrder = Number(row.order.bed) || Number(row.order.bath) || Number(row.order.mats);
        const note = row.note || (row.shouldAnticipate ? 'Anticipare' : hasOrder ? 'Ordinaria' : 'Nessun ordine');
        return `
          <tr class="${row.shouldAnticipate ? 'critical-row' : hasOrder ? 'warn-row' : ''}">
            <td>${esc(row.apartmentLabel || row.apartment)}<br><span style="color: var(--muted); font-size: 12px;">${esc(row.deliveryAddress || '')}</span></td>
            <td class="num">${esc(row.sequence)}</td>
            <td>${fmtDate(row.lastDeliveryDate)}</td>
            <td>${fmtDate(row.scheduledDeliveryDate)}</td>
            <td>${fmtDate(row.suggestedDeliveryDate)}</td>
            <td class="num">${esc(row.order.bed)}</td>
            <td class="num">${esc(row.order.bath)}</td>
            <td class="num">${esc(row.order.mats)}</td>
            <td>${note}</td>
          </tr>
        `;
      }).join('');
      document.getElementById('scheduleTable').innerHTML = scheduleRows || '<tr><td colspan="9" class="empty">Nessuna consegna suggerita.</td></tr>';

      const bookingRows = visibleBookings.map(b => {
        const isDelivery = b.type && b.type !== 'booking';
        const rowThreshold = Number(b.threshold ?? threshold);
        const bedMove = isDelivery ? `+${esc(b.supply.bed)}` : esc(b.consume.bed);
        const bathMove = isDelivery ? `+${esc(b.supply.bath)}` : esc(b.consume.bath);
        const matsMove = isDelivery ? `+${esc(b.supply.mats)}` : esc(b.consume.mats);
        return `
        <tr class="${b.isPast ? 'past-row ' : ''}${isDelivery ? 'warn-row ' : ''}${b.status === 'red' ? 'critical-row' : b.status === 'yellow' ? 'warn-row' : ''}">
          <td>${fmtDate(b.checkIn)}</td>
          <td>${esc(b.apartmentLabel || b.apartment)}</td>
          <td class="num">${isDelivery ? '' : esc(b.guests)}</td>
          <td>${isDelivery ? 'Consegna' : b.extraBed ? 'Si' : ''}</td>
          <td>${esc(b.guestName)}</td>
          <td class="num">${bedMove}</td>
          <td class="num">${bathMove}</td>
          <td class="num">${matsMove}</td>
          <td class="num">${stockCell(b.remaining.bed, rowThreshold)}</td>
          <td class="num">${stockCell(b.remaining.bath, rowThreshold)}</td>
          <td class="num">${stockCell(b.remaining.mats, rowThreshold)}</td>
          <td>${badge(b.status)}</td>
        </tr>
      `}).join('');
      document.getElementById('bookingsTable').innerHTML = bookingRows || '<tr><td colspan="12" class="empty">Nessuna prenotazione riconosciuta.</td></tr>';
      renderPricing(data);
      renderCalendar(data);
      renderCleaning(data);
    }

    addDelivery.addEventListener('click', () => addDeliveryRow({ apartment: apartmentFilter.value || '' }));
    addStaffUser.addEventListener('click', () => addStaffUserRow({ active: '1' }));
    if (addApartment) addApartment.addEventListener('click', () => addApartmentRow());
    configToggle.addEventListener('click', () => {
      configPanel.classList.toggle('active');
      configToggle.textContent = configPanel.classList.contains('active') ? 'Chiudi configurazione' : 'Configura appartamenti';
    });
    staffConfigToggle.addEventListener('click', () => {
      configPanel.classList.add('active');
      configToggle.textContent = 'Chiudi configurazione';
      staffUsersSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
    moduleButtons.forEach(button => {
      button.addEventListener('click', () => setActiveModule(button.dataset.moduleTarget));
    });
    if (calendarPrev) calendarPrev.addEventListener('click', () => {
      calendarMonth = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth() - 1, 1);
      if (lastData) renderCalendar(lastData);
    });
    if (calendarNext) calendarNext.addEventListener('click', () => {
      calendarMonth = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth() + 1, 1);
      if (lastData) renderCalendar(lastData);
    });
    ['bedGuests', 'bathGuests', 'matsStandard', 'matsExtra', 'startBed', 'startBath', 'startMats', 'startDate', 'threshold'].forEach(name => {
      const input = document.querySelector(`[name="${name}"]`);
      if (input) input.addEventListener('input', syncVisibleSettingsToSelectedApartment);
    });
    apartmentFilter.addEventListener('change', () => {
      if (lastData) render(lastData);
    });
    exportDeliveries.addEventListener('click', () => {
      const blob = new Blob([deliveriesToCsv(getDeliveries())], { type: 'text/csv;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `archivio-consegne-biancheria-${new Date().toISOString().slice(0, 10)}.csv`;
      link.click();
      URL.revokeObjectURL(url);
    });
    importDeliveries.addEventListener('click', () => importDeliveriesFile.click());
    importDeliveriesFile.addEventListener('change', async () => {
      const file = importDeliveriesFile.files[0];
      if (!file) return;
      try {
        const text = await file.text();
        const rows = file.name.toLowerCase().endsWith('.csv')
          ? csvToDeliveries(text)
          : (() => {
              const parsed = JSON.parse(text);
              return Array.isArray(parsed) ? parsed : parsed.deliveries;
            })();
        if (!Array.isArray(rows)) throw new Error('Archivio consegne non valido.');
        loadDeliveriesArchive(rows);
      } catch (error) {
        errorBox.textContent = error.message;
        errorBox.style.display = 'block';
      } finally {
        importDeliveriesFile.value = '';
      }
    });
    restoreDeliveriesArchive();
    restoreApartmentsArchive();
    restoreStaffUsersArchive();
    pricingInput.value = JSON.stringify(getStoredPricing());
    if (staffMode) {
      document.body.classList.add('staff-mode');
      setActiveModule('pulizie');
    }

    function supplierText() {
      const rows = lastSchedule.filter(row => Number(row.order.bed) || Number(row.order.bath) || Number(row.order.mats));
      if (!rows.length) return 'Nessuna consegna da ordinare al momento.';
      const lines = ['Buongiorno,', '', 'richiedo consegna biancheria con queste quantita:', ''];
      rows.forEach(row => {
        lines.push(`Consegna ${row.sequence} - ${fmtDate(row.suggestedDeliveryDate)} - ${row.apartmentLabel || row.apartment}`);
        if (row.deliveryAddress) lines.push(`Indirizzo: ${row.deliveryAddress}`);
        lines.push(`Kit letto: ${row.order.bed}`);
        lines.push(`Kit bagno: ${row.order.bath}`);
        lines.push(`Tappeti: ${row.order.mats}`);
        if (row.note) lines.push(`Nota: ${row.note}`);
        lines.push('');
      });
      lines.push('Grazie');
      return lines.join('\n');
    }

    copySupplierText.addEventListener('click', async () => {
      await navigator.clipboard.writeText(supplierText());
      copySupplierText.textContent = 'Copiato';
      setTimeout(() => copySupplierText.textContent = 'Copia testo ordine', 1400);
    });

    emailSupplier.addEventListener('click', () => {
      const subject = encodeURIComponent('Ordine biancheria appartamenti');
      const body = encodeURIComponent(supplierText());
      const to = encodeURIComponent(supplierEmail.value || '');
      window.location.href = `mailto:${to}?subject=${subject}&body=${body}`;
    });

    whatsappSupplier.addEventListener('click', () => {
      const phone = (supplierWhatsapp.value || '').replace(/\D/g, '');
      const text = encodeURIComponent(supplierText());
      const url = phone ? `https://wa.me/${phone}?text=${text}` : `https://wa.me/?text=${text}`;
      window.open(url, '_blank', 'noopener');
    });

    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      errorBox.style.display = 'none';
      noticeBox.style.display = 'none';
      button.disabled = true;
      refreshButton.disabled = true;
      button.textContent = 'Aggiorno...';
      refreshButton.textContent = 'Aggiorno...';
      try {
        serializeDeliveries();
        serializeStaffUsers();
        syncVisibleSettingsToSelectedApartment();
        serializeApartments();
        savePricingArchive();
        saveDeliveriesArchive();
        saveStaffUsersArchive();
        saveApartmentsArchive();
        const body = new FormData(form);
        const response = await fetch('/api/upload', { method: 'POST', body });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Errore durante il caricamento.');
        render(data);
      } catch (error) {
        errorBox.textContent = error.message;
        errorBox.style.display = 'block';
      } finally {
        button.disabled = false;
        refreshButton.disabled = false;
        button.textContent = 'Carica e aggiorna';
        refreshButton.textContent = 'Aggiorna calcoli';
      }
    });
  </script>
</body>
</html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def send_json(self, payload, code=200):
        raw = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        clean_path = self.path.split("?", 1)[0]
        if clean_path.startswith("/api/"):
            self.send_error(404)
            return
        raw = INDEX_HTML.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self):
        if self.path != "/api/upload":
            self.send_error(404)
            return
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type"),
                },
            )
            upload = form["file"]
            raw = upload.file.read()
            if not raw:
                self.send_json({"error": "Il file caricato e vuoto."}, 400)
                return
            filename = str(getattr(upload, "filename", "") or "")
            if not filename.lower().endswith(".xlsx"):
                self.send_json({"error": "Formato file non supportato: carica un file Excel .xlsx."}, 400)
                return
            if not raw.startswith(b"PK"):
                self.send_json({"error": "Il file non sembra un .xlsx valido. Se e un .xls o viene da Numbers, esportalo come Excel .xlsx."}, 400)
                return
            settings = {
                "threshold": form_number(form, "threshold", 2),
                "bedGuests": form_number(form, "bedGuests", 2),
                "bathGuests": form_number(form, "bathGuests", 1),
                "matsStandard": form_number(form, "matsStandard", 1),
                "matsExtra": form_number(form, "matsExtra", 0),
                "startBed": form_number(form, "startBed", 0),
                "startBath": form_number(form, "startBath", 0),
                "startMats": form_number(form, "startMats", 0),
                "startDate": str(form.getvalue("startDate") or "").strip(),
            }
            try:
                settings["apartments"] = json.loads(form.getvalue("apartments") or "[]")
            except json.JSONDecodeError:
                settings["apartments"] = []
            try:
                settings["pricing"] = json.loads(form.getvalue("pricing") or "[]")
            except json.JSONDecodeError:
                settings["pricing"] = []
            try:
                deliveries = json.loads(form.getvalue("deliveries") or "[]")
            except json.JSONDecodeError:
                deliveries = []
            self.send_json(parse_workbook(raw, settings=settings, deliveries=deliveries))
        except Exception as exc:
            traceback.print_exc()
            self.send_json({"error": html.escape(str(exc))}, 500)


if __name__ == "__main__":
    shown_host = "localhost" if HOST in {"127.0.0.1", "0.0.0.0"} else HOST
    print(f"Gestione appartamenti attiva su http://{shown_host}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
