import base64
import io
from datetime import date
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

from odoo import _, fields, models
from odoo.exceptions import UserError


XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("", XLSX_NS)


PAPER_XLSX_CONFIG = {
    "401": {
        "sheet_name": "401 申報書",
        "sheet_path": "xl/worksheets/sheet1.xml",
        "digit_fields": {
            "vat_no": ["D4", "G4", "I4", "K4", "M4", "O4", "Q4", "S4"],
            "tax_id_9": ["D6", "F6", "H6", "J6", "L6", "N6", "P6", "R6", "T6"],
        },
        "text_fields": {
            "company_name": "D5",
            "roc_year": "AB6",
            "period_month": "AE6",
            "responsible_name": "D7",
            "company_address": "X7",
            "invoice_count": "AV8",
            "sales_1_amount": "W10",
            "sales_2_tax": "AA10",
            "sales_5_amount": "W11",
            "sales_6_tax": "AA11",
            "sales_7_zero": "AG11",
            "sales_9_amount": "W12",
            "sales_10_tax": "AA12",
            "sales_13_amount": "W13",
            "sales_14_tax": "AA13",
            "sales_15_zero": "AG13",
            "sales_17_amount": "W14",
            "sales_18_tax": "AA14",
            "sales_19_zero": "AG14",
            "sales_21_amount": "W15",
            "sales_22_tax": "AA15",
            "sales_23_zero": "AG15",
            "sales_25_total": "W16",
            "sales_27_fixed_asset": "AD16",
            "purchase_28_amount_goods": "Y20",
            "purchase_29_tax_goods": "AC20",
            "purchase_30_amount_asset": "Y21",
            "purchase_31_tax_asset": "AC21",
            "purchase_32_amount_goods": "Y22",
            "purchase_33_tax_goods": "AC22",
            "purchase_34_amount_asset": "Y23",
            "purchase_35_tax_asset": "AC23",
            "purchase_36_amount_goods": "Y24",
            "purchase_37_tax_goods": "AC24",
            "purchase_38_amount_asset": "Y25",
            "purchase_39_tax_asset": "AC25",
            "purchase_78_amount_goods": "Y26",
            "purchase_79_tax_goods": "AC26",
            "purchase_80_amount_asset": "Y27",
            "purchase_81_tax_asset": "AC27",
            "purchase_40_amount_goods": "Y28",
            "purchase_41_tax_goods": "AC28",
            "purchase_42_amount_asset": "Y29",
            "purchase_43_tax_asset": "AC29",
            "purchase_44_amount_goods": "Y30",
            "purchase_45_tax_goods": "AC30",
            "purchase_46_amount_asset": "Y31",
            "purchase_47_tax_asset": "AC31",
            "purchase_48_total_goods": "Y32",
            "purchase_49_total_asset": "Y33",
            "import_tax_exempt_goods": "W34",
            "purchase_foreign_services": "W35",
            "item_101_output_tax": "AV9",
            "item_107_input_tax_total": "AV10",
            "item_108_previous_credit": "AV11",
            "item_110_input_tax_subtotal": "AV12",
            "item_111_net_tax_payable": "AV13",
            "item_112_net_tax_credit": "AV14",
            "item_113_refund_limit": "AV15",
            "item_114_refund_amount": "AV16",
            "item_115_accumulated_credit": "AV17",
            "filing_date": "AM31",
            "self_filer_name": "AM34",
            "self_filer_idno": "AN34",
            "self_filer_phone": "AT34",
            "self_filer_reg_no": "AW34",
            "agent_name": "AM35",
            "agent_idno": "AN35",
            "agent_phone": "AT35",
            "agent_reg_no": "AW35",
        },
        "note": "",
    },
    "403": {
        "sheet_name": "403 申報書",
        "sheet_path": "xl/worksheets/sheet3.xml",
        "digit_fields": {
            "vat_no": ["D4", "G4", "I4", "K4", "M4", "O4", "Q4", "S4"],
            "tax_id_9": ["D6", "F6", "H6", "J6", "L6", "N6", "P6", "R6", "T6"],
        },
        "text_fields": {
            "company_name": "D5",
            "roc_year": "AC6",
            "period_month": "AE6",
            "responsible_name": "D7",
            "company_address": "X7",
            "invoice_count": "AV7",
            "sales_1_amount": "W10",
            "sales_2_tax": "AC10",
            "sales_3_zero": "AK10",
            "sales_4_exempt": "AT10",
            "sales_5_amount": "W11",
            "sales_6_tax": "AC11",
            "sales_7_zero": "AK11",
            "sales_8_exempt": "AT11",
            "sales_9_amount": "W12",
            "sales_10_tax": "AC12",
            "sales_11_zero": "AK12",
            "sales_12_exempt": "AT12",
            "sales_13_amount": "W13",
            "sales_14_tax": "AC13",
            "sales_15_zero": "AK13",
            "sales_16_exempt": "AT13",
            "sales_17_amount": "W14",
            "sales_18_tax": "AC14",
            "sales_19_zero": "AK14",
            "sales_20_exempt": "AT14",
            "sales_21_amount": "W15",
            "sales_22_tax": "AC15",
            "sales_23_zero": "AK15",
            "sales_24_exempt": "AT15",
            "special_52_amount": "W17",
            "special_53_tax": "AC17",
            "special_54_amount": "W18",
            "special_55_tax": "AC18",
            "special_84_amount": "W19",
            "special_85_tax": "AC19",
            "special_56_amount": "W20",
            "special_57_tax": "AC20",
            "special_60_amount": "W21",
            "special_61_tax": "AC21",
            "special_62_amount": "W22",
            "special_63_amount": "W23",
            "special_64_tax": "AC23",
            "special_65_amount": "W24",
            "special_66_tax": "AC24",
            "sales_total_403": "W25",
            "sales_land_amount_26": "AF25",
            "purchase_28_amount_goods": "Y29",
            "purchase_29_tax_goods": "AC29",
            "purchase_30_amount_asset": "Y30",
            "purchase_31_tax_asset": "AC30",
            "purchase_32_amount_goods": "Y31",
            "purchase_33_tax_goods": "AC31",
            "purchase_34_amount_asset": "Y32",
            "purchase_35_tax_asset": "AC32",
            "purchase_36_amount_goods": "Y33",
            "purchase_37_tax_goods": "AC33",
            "purchase_38_amount_asset": "Y34",
            "purchase_39_tax_asset": "AC34",
            "purchase_78_amount_goods": "Y35",
            "purchase_79_tax_goods": "AC35",
            "purchase_80_amount_asset": "Y36",
            "purchase_81_tax_asset": "AC36",
            "purchase_40_amount_goods": "Y37",
            "purchase_41_tax_goods": "AC37",
            "purchase_42_amount_asset": "Y38",
            "purchase_43_tax_asset": "AC38",
            "purchase_44_amount_goods": "Y39",
            "purchase_45_tax_goods": "AC39",
            "purchase_46_amount_asset": "Y40",
            "purchase_47_tax_asset": "AC40",
            "purchase_48_total_goods": "Y41",
            "purchase_49_total_asset": "Y42",
            "nondeductible_ratio_50": "Z43",
            "input_tax_deductible_51": "Z45",
            "item_101_output_tax_total": "AT17",
            "item_103_foreign_services": "AT18",
            "item_104_special_tax": "AT19",
            "item_105_adjustment_due": "AT20",
            "item_106_subtotal": "AT21",
            "item_107_input_tax_total": "AT22",
            "item_108_previous_credit": "AT23",
            "item_109_adjustment_refund": "AT24",
            "item_110_input_tax_subtotal": "AT25",
            "item_111_net_tax_payable": "AT26",
            "item_112_net_tax_credit": "AT27",
            "item_113_refund_limit": "AT28",
            "item_114_refund_amount": "AT29",
            "item_115_accumulated_credit": "AT30",
            "import_tax_exempt_goods_73": "Q47",
            "foreign_services_amount_74": "Q49",
            "foreign_services_tax_75": "AA49",
            "foreign_services_payable_76": "AG49",
            "filing_date": "AM46",
            "self_filer_name": "AL48",
            "self_filer_idno": "AN48",
            "self_filer_phone": "AT48",
            "self_filer_reg_no": "AW48",
            "agent_name": "AL49",
            "agent_idno": "AN49",
            "agent_phone": "AT49",
            "agent_reg_no": "AW49",
        },
        "note": "403申報書已依官方欄位帶入一般稅額區、特種稅額區、比例扣抵進項稅額與右側申報摘要；如有購買國外勞務或中途歇業/年底調整，請再人工確認。",
    },
    "404": {
        "sheet_name": "404 申報書",
        "sheet_path": "xl/worksheets/sheet5.xml",
        "digit_fields": {
            "vat_no": ["E6", "H6", "J6", "L6", "N6", "P6", "R6", "T6"],
            "tax_id_9": ["E8", "G8", "I8", "K8", "M8", "O8", "Q8", "S8", "U8"],
        },
        "text_fields": {
            "roc_year": "Y5",
            "period_month": "AA5",
            "company_name": "E7",
            "responsible_name": "Y6",
            "company_address": "Y7",
            "net_tax_payable": "AH24",
            "filing_date": "H30",
            "self_filer_name": "E32",
            "self_filer_idno": "N32",
            "self_filer_phone": "W32",
            "self_filer_reg_no": "Y32",
            "agent_name": "E33",
            "agent_idno": "N33",
            "agent_phone": "W33",
            "agent_reg_no": "Y33",
        },
        "note": "404申報書目前僅自動帶入表頭、申報人資料與本期應實繳稅額；特種業別明細仍需人工填寫。",
    },
}


def _xlsx_tag(tag_name):
    return f"{{{XLSX_NS}}}{tag_name}"


def _digits_only(value):
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _upper_alnum(value):
    return "".join(ch for ch in str(value or "").upper() if ch.isalnum())


def _zfill_digits(value, length):
    return _digits_only(value).zfill(length)[:length]


def _format_period_month(value):
    month = int(value or 0)
    if month < 1 or month > 12:
        return ""
    if month % 2 == 0:
        return f"{month - 1}—{month}"
    return str(month)


def _rpad(value, length):
    value = str(value or "")
    return (value + (" " * length))[:length]


def _vat8_or_blank(vat):
    digits = _digits_only((vat or "").replace("TW", ""))
    return digits if len(digits) == 8 else " " * 8


def _clean(value):
    if value in (None, False, ""):
        return ""
    text = " ".join(str(value).split())
    return text.replace("\r", " ").replace("\n", " ")


class TwVatFilingWizard(models.TransientModel):
    _name = "tw.vat.filing.wizard"
    _description = "臺灣營業稅申報"

    company_id = fields.Many2one("res.company", required=True, default=lambda self: self.env.company, string="公司")
    year_roc = fields.Integer(string="年度(民國)", required=True, default=lambda self: self._default_filing_period()[0])
    month = fields.Integer(string="月份(期末月)", required=True, default=lambda self: self._default_filing_period()[1])
    filing_code = fields.Selection([("1", "1"), ("5", "5")], string="申報代號", required=True, default="1")
    total_pay_code = fields.Selection([("0", "0"), ("1", "1"), ("2", "2")], string="總繳代號", required=True, default="0")
    tax_rate_percent = fields.Float(string="一般稅額徵收率(%)", default=5.0)
    special_tax_rate_percent = fields.Float(string="特種稅額稅率(%)", default=0.0)
    paper_form_type = fields.Selection(
        [("401", "401"), ("403", "403"), ("404", "404")],
        string="紙本申報書類型",
        default="401",
    )
    export_zip = fields.Binary(string="申報匯出ZIP", readonly=True)
    export_zip_name = fields.Char(string="匯出檔名", readonly=True)
    check_report = fields.Text(string="匯出檢核報表", readonly=True)
    paper_xlsx = fields.Binary(string="進銷項核對表Excel", readonly=True)
    paper_xlsx_name = fields.Char(string="Excel檔名", readonly=True)
    paper_check_report = fields.Text(string="Excel匯出檢核報表", readonly=True)
    def _default_filing_period(self):
        today = fields.Date.context_today(self)
        year = today.year
        month = today.month
        if month % 2 == 1:
            month -= 1
        else:
            month -= 2
        if month <= 0:
            year -= 1
            month += 12
        return year - 1911, month

    def _period_range(self):
        year = int(self.year_roc) + 1911
        month = int(self.month)
        if month < 1 or month > 12:
            raise UserError(_("月份需為1~12"))
        date_from = date(year, month, 1)
        date_to = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
        return date_from, date_to

    def _get_moves(self, date_from, date_to):
        domain = [
            ("company_id", "=", self.company_id.id),
            ("state", "=", "posted"),
            ("invoice_date", ">=", date_from),
            ("invoice_date", "<", date_to),
            ("move_type", "in", ("out_invoice", "out_refund", "in_invoice", "in_refund")),
        ]
        if "tw_blr_skip_export" in self.env["account.move"]._fields:
            domain.append(("tw_blr_skip_export", "=", False))
        return self.env["account.move"].search(domain)

    def _compute_amounts_for_export(self, move, lines=None, deduct_code=None):
        tax_type = (move.tw_tax_type or "")[:1]
        deduct_code = deduct_code or move._blr_get_export_deduct_code()
        if lines is None:
            amount_untaxed = int(round(move.amount_untaxed))
            amount_total = int(round(move.amount_total))
            tax_amount = int(round(move.amount_tax))
        else:
            untaxed_total = sum((line.price_subtotal or 0.0) for line in lines)
            amount_total_raw = sum((line.price_total or 0.0) for line in lines)
            amount_untaxed = int(round(untaxed_total))
            amount_total = int(round(amount_total_raw))
            tax_amount = amount_total - amount_untaxed
        if tax_type == "1" and deduct_code in ("3", "4") and tax_amount == 0:
            rate = float(self.tax_rate_percent or 0.0) / 100.0
            if rate > 0:
                amount_untaxed = int(round(amount_total / (1.0 + rate)))
                tax_amount = amount_total - amount_untaxed
        return amount_untaxed, tax_amount

    def _get_split_deduct_codes(self, move):
        tax_type = (move.tw_tax_type or "")[:1]
        if tax_type in ("2", "3"):
            return "3", "4"
        deduct_code = move._blr_get_export_deduct_code()
        if deduct_code in ("1", "2"):
            return "1", "2"
        if deduct_code in ("3", "4"):
            return "3", "4"
        return False, False

    def _build_purchase_export_entries(self, move):
        invoice_lines = move._blr_relevant_invoice_lines()
        goods_lines = invoice_lines.filtered(
            lambda line: getattr(line.account_id, "account_type", "") != "asset_fixed"
        )
        asset_lines = invoice_lines.filtered(
            lambda line: getattr(line.account_id, "account_type", "") == "asset_fixed"
        )
        goods_code, asset_code = self._get_split_deduct_codes(move)
        entry_specs = [
            ("goods", goods_lines, goods_code, "進貨及費用"),
            ("asset", asset_lines, asset_code, "固定資產"),
        ]
        active_specs = [spec for spec in entry_specs if spec[1]]
        entries = []
        for entry_kind, lines, deduct_code, label in active_specs:
            amount_untaxed, tax_amount = self._compute_amounts_for_export(move, lines=lines, deduct_code=deduct_code)
            if not amount_untaxed and not tax_amount:
                raw_amount = sum((line.price_subtotal or 0.0) + (line.price_total or 0.0) for line in lines)
                if not raw_amount:
                    continue
            display_name = move.display_name if len(active_specs) == 1 else _("%s（%s）") % (move.display_name, label)
            entries.append({
                "move": move,
                "display_name": display_name,
                "format_code": move.tw_blr_format_code or "",
                "tax_type": (move.tw_tax_type or "")[:1],
                "deduct_code": deduct_code or " ",
                "amount_untaxed": amount_untaxed,
                "tax_amount": tax_amount,
                "entry_kind": entry_kind,
            })
        if entries:
            return entries

        amount_untaxed, tax_amount = self._compute_amounts_for_export(move)
        if not amount_untaxed and not tax_amount and not int(round(move.amount_total or 0.0)):
            return []
        deduct_code = move._blr_get_export_deduct_code() or " "
        entry_kind = "asset" if deduct_code in ("2", "4") else "goods"
        return [{
            "move": move,
            "display_name": move.display_name,
            "format_code": move.tw_blr_format_code or "",
            "tax_type": (move.tw_tax_type or "")[:1],
            "deduct_code": deduct_code,
            "amount_untaxed": amount_untaxed,
            "tax_amount": tax_amount,
            "entry_kind": entry_kind,
        }]

    def _build_export_entries(self, move):
        if move.move_type in ("in_invoice", "in_refund"):
            return self._build_purchase_export_entries(move)
        amount_untaxed, tax_amount = self._compute_amounts_for_export(move)
        return [{
            "move": move,
            "display_name": move.display_name,
            "format_code": move.tw_blr_format_code or "",
            "tax_type": (move.tw_tax_type or "")[:1],
            "deduct_code": move._blr_get_export_deduct_code(),
            "amount_untaxed": amount_untaxed,
            "tax_amount": tax_amount,
            "entry_kind": "sale",
        }]

    def _select_identifier(self, move, format_code):
        track = _clean(move.tw_invoice_track).upper()
        invoice_number = _zfill_digits(move.tw_invoice_number, 8)
        other_voucher = _upper_alnum(move.tw_other_voucher_no)
        utility_carrier = _upper_alnum(move.tw_utility_carrier_no)
        customs_no = _upper_alnum(move.tw_customs_pay_no)
        has_invoice = bool(track or _digits_only(invoice_number) != "00000000")
        has_other = bool(other_voucher)
        has_utility = bool(utility_carrier)

        if format_code in ("37", "38"):
            raise UserError(_("目前版本未開放特種稅額格式代號%s") % format_code)

        if format_code in ("28", "29"):
            if len(customs_no) != 14 or has_invoice or has_other or has_utility:
                raise UserError(_("格式代號%s僅能填海關代徵營業稅繳納證號碼(14碼)") % format_code)
            return "customs", customs_no

        if format_code in ("25", "35"):
            if has_utility and has_invoice:
                raise UserError(_("格式代號%s不可同時填公用事業載具流水號與發票字軌號碼") % format_code)
            if utility_carrier.startswith("BB") and len(utility_carrier) == 10:
                return "utility", utility_carrier
            if len(track) == 2 and _digits_only(invoice_number) != "00000000":
                return "invoice", track + invoice_number
            raise UserError(_("格式代號%s需擇一填發票字軌/號碼或公用事業載具流水號(BB+8碼)") % format_code)

        if format_code in ("22", "24", "27", "32", "34", "36"):
            if has_other and has_invoice:
                raise UserError(_("格式代號%s不可同時填其他憑證號碼與發票字軌號碼") % format_code)
            if other_voucher and len(other_voucher) == 10:
                return "other", other_voucher
            if len(track) == 2 and _digits_only(invoice_number) != "00000000":
                return "invoice", track + invoice_number
            raise UserError(_("格式代號%s需擇一填發票字軌/號碼或其他憑證號碼(10碼)") % format_code)

        if len(track) == 2 and _digits_only(invoice_number) != "00000000":
            return "invoice", track + invoice_number
        raise UserError(_("格式代號%s需填發票字軌(2碼)+發票號碼(8碼)") % format_code)

    def _txt_line_81(self, move, sequence, year_roc, month, entry=None):
        entry = entry or {}
        display_name = entry.get("display_name") or move.display_name
        format_code = entry.get("format_code") or move.tw_blr_format_code
        if not format_code or len(format_code) != 2:
            raise UserError(_("單據 %s 缺少申報格式代號") % display_name)

        tax_id_9 = _rpad(_digits_only(self.company_id.tw_tax_id_9), 9)
        if len(_digits_only(tax_id_9)) != 9:
            raise UserError(_("公司稅籍編號需為9碼"))

        sequence_7 = _zfill_digits(sequence, 7)
        year_3 = _zfill_digits(year_roc, 3)
        month_2 = _zfill_digits(month, 2)
        company_vat_8 = _vat8_or_blank(self.company_id.vat)
        partner_vat_8 = _vat8_or_blank(move.partner_id.vat)

        if len(_digits_only(company_vat_8)) != 8:
            raise UserError(_("公司統一編號需為8碼"))

        if move.move_type in ("out_invoice", "out_refund"):
            buyer_vat_8 = partner_vat_8 if partner_vat_8 else " " * 8
            seller_vat_8 = company_vat_8
        else:
            buyer_vat_8 = company_vat_8
            seller_vat_8 = partner_vat_8 if partner_vat_8 else " " * 8

        identifier_kind, identifier = self._select_identifier(move, format_code)
        track_2 = "  "
        invoice_8 = "00000000"
        if identifier_kind == "invoice":
            track_2 = identifier[:2]
            invoice_8 = identifier[2:]
        elif identifier_kind == "other":
            track_2 = "OT"
            invoice_8 = identifier[:8]
        elif identifier_kind == "utility":
            track_2 = "BB"
            invoice_8 = identifier[2:]
        elif identifier_kind == "customs":
            track_2 = "CU"
            invoice_8 = identifier[:8]

        amount_untaxed = entry.get("amount_untaxed")
        tax_amount = entry.get("tax_amount")
        tax_type = entry.get("tax_type") or (move.tw_tax_type or "")[:1]
        deduct_code = entry.get("deduct_code")
        if amount_untaxed is None or tax_amount is None:
            amount_untaxed, tax_amount = self._compute_amounts_for_export(move)
        if deduct_code in (None, ""):
            deduct_code = move._blr_get_export_deduct_code()
        if move.move_type in ("out_invoice", "out_refund"):
            deduct_code = " "

        if tax_type not in ("1", "2", "3"):
            raise UserError(_("單據 %s 缺少有效的課稅別") % display_name)
        if move.move_type in ("in_invoice", "in_refund") and deduct_code not in ("1", "2", "3", "4"):
            raise UserError(_("單據 %s 缺少有效的抵扣代號") % display_name)

        if tax_type in ("2", "3") and tax_amount != 0:
            raise UserError(_("單據 %s 課稅別為零稅率或免稅時，稅額需為0") % display_name)
        if tax_type == "1":
            expected_tax = int(round(amount_untaxed * (self.tax_rate_percent / 100.0)))
            if abs(expected_tax - tax_amount) > 1:
                raise UserError(
                    _("單據 %s 稅額檢核失敗：未稅=%s 推算稅額約=%s 實際稅額=%s")
                    % (display_name, amount_untaxed, expected_tax, tax_amount)
                )
        if tax_type in ("2", "3") and deduct_code in ("1", "2"):
            raise UserError(_("單據 %s 課稅別為零稅率或免稅時，抵扣代號不可為1或2") % display_name)

        line = ""
        line += _rpad(format_code, 2)
        line += _rpad(tax_id_9, 9)
        line += _rpad(sequence_7, 7)
        line += year_3
        line += month_2
        line += _rpad(buyer_vat_8, 8)
        line += _rpad(seller_vat_8, 8)
        line += _rpad(track_2, 2)
        line += _rpad(invoice_8, 8)
        line += _rpad(_zfill_digits(amount_untaxed, 12), 12)
        line += _rpad(tax_type, 1)
        line += _rpad(_zfill_digits(tax_amount, 10), 10)
        line += _rpad(deduct_code, 1)
        line += " " * 8

        if len(line) != 81:
            raise UserError(_("TXT單行必須81字元，實際=%s") % len(line))
        return line

    def _tet_u_line(self, totals):
        fields_data = ["0"] * 112
        fields_data[0] = "1"
        fields_data[1] = "00000001"
        fields_data[2] = _zfill_digits((self.company_id.vat or "").replace("TW", ""), 8)
        fields_data[3] = _zfill_digits(self.year_roc, 3) + _zfill_digits(self.month, 2)
        fields_data[4] = self.filing_code
        fields_data[5] = _rpad(_digits_only(self.company_id.tw_tax_id_9), 9)
        fields_data[6] = self.total_pay_code
        fields_data[7] = _zfill_digits(totals.get("invoice_count", 0), 10)
        fields_data[8] = _zfill_digits(totals.get("taxable_sales", 0), 12)
        fields_data[24] = _zfill_digits(totals.get("output_tax", 0), 10)
        fields_data[72] = _zfill_digits(totals.get("input_tax_deductible", 0), 10)
        fields_data = [_clean(value) for value in fields_data]
        return "|".join(fields_data)
    def _prepare_export_run(self):
        company = self.company_id
        if not company.vat:
            raise UserError(_("請先在公司資料填寫統一編號(8碼)"))
        if not company.tw_tax_id_9 or len(_digits_only(company.tw_tax_id_9)) != 9:
            raise UserError(_("請先在公司資料填寫稅籍編號(9碼)"))

        date_from, date_to = self._period_range()
        moves = self._get_moves(date_from, date_to)
        year_roc = int(self.year_roc)
        month = int(self.month)
        rate = float(self.tax_rate_percent or 0.0) / 100.0
        errors = []
        exported = []
        txt_lines = []
        sequence = 1
        invoice_count = 0
        taxable_sales = 0
        zero_rate_sales = 0
        exempt_sales = 0
        output_tax = 0
        input_tax_deductible = 0
        input_tax_deductible_goods = 0
        input_tax_deductible_asset = 0

        for move in moves.sorted(lambda record: (record.invoice_date or record.date, record.name)):
            if not move.tw_blr_format_code:
                exported.append((move.display_name, "跳過：缺申報格式代號"))
                continue
            try:
                move._blr_validate_fields_raise()
                entries = self._build_export_entries(move)
                if not entries:
                    exported.append((move.display_name, "跳過：沒有可匯出的有效資料"))
                    continue
                for entry in entries:
                    txt_lines.append(self._txt_line_81(move, sequence, year_roc, month, entry=entry))
                    sequence += 1
                    invoice_count += 1
                    amount_untaxed = int(entry.get("amount_untaxed", 0) or 0)
                    tax_amount = int(entry.get("tax_amount", 0) or 0)
                    tax_type = entry.get("tax_type") or ""
                    deduct_code = entry.get("deduct_code") or " "
                    if move.move_type in ("out_invoice", "out_refund"):
                        if tax_type == "1":
                            taxable_sales += amount_untaxed
                            output_tax += tax_amount
                        elif tax_type == "2":
                            zero_rate_sales += amount_untaxed
                        elif tax_type == "3":
                            exempt_sales += amount_untaxed
                    elif deduct_code in ("1", "2"):
                        input_tax_deductible += tax_amount
                        if deduct_code == "1":
                            input_tax_deductible_goods += tax_amount
                        else:
                            input_tax_deductible_asset += tax_amount
                status = "OK" if len(entries) == 1 else _("OK，已依進貨及費用/固定資產拆分%s筆") % len(entries)
                exported.append((move.display_name, status))
            except Exception as exc:
                errors.append(f"{move.display_name}: {exc}")

        input_tax_subtotal = input_tax_deductible
        refund_limit = max(int(round(zero_rate_sales * rate)), 0)
        net_tax_payable = max(output_tax - input_tax_subtotal, 0)
        net_tax_credit = max(input_tax_subtotal - output_tax, 0)
        refund_amount = min(net_tax_credit, refund_limit)
        accumulated_credit = max(net_tax_credit - refund_amount, 0)
        totals = {
            "invoice_count": invoice_count,
            "taxable_sales": taxable_sales,
            "zero_rate_sales": zero_rate_sales,
            "exempt_sales": exempt_sales,
            "sales_total_401": taxable_sales + zero_rate_sales,
            "sales_total_403": taxable_sales + zero_rate_sales + exempt_sales,
            "output_tax": output_tax,
            "input_tax_deductible": input_tax_deductible,
            "input_tax_deductible_goods": input_tax_deductible_goods,
            "input_tax_deductible_asset": input_tax_deductible_asset,
            "previous_credit": 0,
            "input_tax_subtotal": input_tax_subtotal,
            "refund_limit": refund_limit,
            "refund_amount": refund_amount,
            "net_tax_payable": net_tax_payable,
            "net_tax_credit": net_tax_credit,
            "accumulated_credit": accumulated_credit,
        }
        return {
            "company": company,
            "moves": moves,
            "errors": errors,
            "exported": exported,
            "txt_lines": txt_lines,
            "totals": totals,
        }

    def _init_401_paper_totals(self):
        keys = (
            "sales_1_amount",
            "sales_2_tax",
            "sales_3_zero",
            "sales_5_amount",
            "sales_6_tax",
            "sales_7_zero",
            "sales_9_amount",
            "sales_10_tax",
            "sales_11_zero",
            "sales_13_amount",
            "sales_14_tax",
            "sales_15_zero",
            "sales_17_amount",
            "sales_18_tax",
            "sales_19_zero",
            "sales_21_amount",
            "sales_22_tax",
            "sales_23_zero",
            "sales_25_total",
            "sales_27_fixed_asset",
            "purchase_28_amount_goods",
            "purchase_29_tax_goods",
            "purchase_30_amount_asset",
            "purchase_31_tax_asset",
            "purchase_32_amount_goods",
            "purchase_33_tax_goods",
            "purchase_34_amount_asset",
            "purchase_35_tax_asset",
            "purchase_36_amount_goods",
            "purchase_37_tax_goods",
            "purchase_38_amount_asset",
            "purchase_39_tax_asset",
            "purchase_40_amount_goods",
            "purchase_41_tax_goods",
            "purchase_42_amount_asset",
            "purchase_43_tax_asset",
            "purchase_44_amount_goods",
            "purchase_45_tax_goods",
            "purchase_46_amount_asset",
            "purchase_47_tax_asset",
            "purchase_48_total_goods",
            "purchase_49_total_asset",
            "purchase_78_amount_goods",
            "purchase_79_tax_goods",
            "purchase_80_amount_asset",
            "purchase_81_tax_asset",
            "import_tax_exempt_goods",
            "purchase_foreign_services",
        )
        return {key: 0 for key in keys}

    def _apply_401_sale_entry(self, detail, entry):
        format_code = entry.get("format_code") or ""
        tax_type = entry.get("tax_type") or ""
        amount_untaxed = abs(int(entry.get("amount_untaxed", 0) or 0))
        tax_amount = abs(int(entry.get("tax_amount", 0) or 0))
        if format_code == "31":
            if tax_type == "1":
                detail["sales_1_amount"] += amount_untaxed
                detail["sales_2_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_3_zero"] += amount_untaxed
        elif format_code == "35":
            if tax_type == "1":
                detail["sales_5_amount"] += amount_untaxed
                detail["sales_6_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_7_zero"] += amount_untaxed
        elif format_code == "32":
            if tax_type == "1":
                detail["sales_9_amount"] += amount_untaxed
                detail["sales_10_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_11_zero"] += amount_untaxed
        elif format_code == "36":
            if tax_type == "1":
                detail["sales_13_amount"] += amount_untaxed
                detail["sales_14_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_15_zero"] += amount_untaxed
        elif format_code in ("33", "34"):
            if tax_type == "1":
                detail["sales_17_amount"] += amount_untaxed
                detail["sales_18_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_19_zero"] += amount_untaxed

    def _apply_401_purchase_entry(self, detail, entry):
        format_code = entry.get("format_code") or ""
        entry_kind = entry.get("entry_kind") or "goods"
        amount_untaxed = abs(int(entry.get("amount_untaxed", 0) or 0))
        tax_amount = abs(int(entry.get("tax_amount", 0) or 0))
        deduct_code = entry.get("deduct_code") or ""
        is_asset = entry_kind == "asset"
        total_key = "purchase_49_total_asset" if is_asset else "purchase_48_total_goods"
        sign = -1 if format_code in ("23", "24", "29") else 1
        detail[total_key] += sign * amount_untaxed

        if deduct_code not in ("1", "2"):
            return

        if format_code in ("21", "26"):
            amount_key = "purchase_30_amount_asset" if is_asset else "purchase_28_amount_goods"
            tax_key = "purchase_31_tax_asset" if is_asset else "purchase_29_tax_goods"
        elif format_code == "25":
            amount_key = "purchase_34_amount_asset" if is_asset else "purchase_32_amount_goods"
            tax_key = "purchase_35_tax_asset" if is_asset else "purchase_33_tax_goods"
        elif format_code in ("22", "27"):
            amount_key = "purchase_38_amount_asset" if is_asset else "purchase_36_amount_goods"
            tax_key = "purchase_39_tax_asset" if is_asset else "purchase_37_tax_goods"
        elif format_code == "28":
            amount_key = "purchase_80_amount_asset" if is_asset else "purchase_78_amount_goods"
            tax_key = "purchase_81_tax_asset" if is_asset else "purchase_79_tax_goods"
        elif format_code in ("23", "24", "29"):
            amount_key = "purchase_42_amount_asset" if is_asset else "purchase_40_amount_goods"
            tax_key = "purchase_43_tax_asset" if is_asset else "purchase_41_tax_goods"
        else:
            return

        detail[amount_key] += amount_untaxed
        detail[tax_key] += tax_amount

    def _finalize_401_paper_totals(self, detail):
        detail["sales_21_amount"] = (
            detail["sales_1_amount"]
            + detail["sales_5_amount"]
            + detail["sales_9_amount"]
            + detail["sales_13_amount"]
            - detail["sales_17_amount"]
        )
        detail["sales_22_tax"] = (
            detail["sales_2_tax"]
            + detail["sales_6_tax"]
            + detail["sales_10_tax"]
            + detail["sales_14_tax"]
            - detail["sales_18_tax"]
        )
        detail["sales_23_zero"] = (
            detail["sales_3_zero"]
            + detail["sales_7_zero"]
            + detail["sales_11_zero"]
            + detail["sales_15_zero"]
            - detail["sales_19_zero"]
        )
        detail["sales_25_total"] = detail["sales_21_amount"] + detail["sales_23_zero"]
        detail["sales_27_fixed_asset"] = 0
        detail["purchase_44_amount_goods"] = (
            detail["purchase_28_amount_goods"]
            + detail["purchase_32_amount_goods"]
            + detail["purchase_36_amount_goods"]
            + detail["purchase_78_amount_goods"]
            - detail["purchase_40_amount_goods"]
        )
        detail["purchase_45_tax_goods"] = (
            detail["purchase_29_tax_goods"]
            + detail["purchase_33_tax_goods"]
            + detail["purchase_37_tax_goods"]
            + detail["purchase_79_tax_goods"]
            - detail["purchase_41_tax_goods"]
        )
        detail["purchase_46_amount_asset"] = (
            detail["purchase_30_amount_asset"]
            + detail["purchase_34_amount_asset"]
            + detail["purchase_38_amount_asset"]
            + detail["purchase_80_amount_asset"]
            - detail["purchase_42_amount_asset"]
        )
        detail["purchase_47_tax_asset"] = (
            detail["purchase_31_tax_asset"]
            + detail["purchase_35_tax_asset"]
            + detail["purchase_39_tax_asset"]
            + detail["purchase_81_tax_asset"]
            - detail["purchase_43_tax_asset"]
        )

    def _init_403_paper_totals(self):
        keys = (
            "sales_1_amount",
            "sales_2_tax",
            "sales_3_zero",
            "sales_4_exempt",
            "sales_5_amount",
            "sales_6_tax",
            "sales_7_zero",
            "sales_8_exempt",
            "sales_9_amount",
            "sales_10_tax",
            "sales_11_zero",
            "sales_12_exempt",
            "sales_13_amount",
            "sales_14_tax",
            "sales_15_zero",
            "sales_16_exempt",
            "sales_17_amount",
            "sales_18_tax",
            "sales_19_zero",
            "sales_20_exempt",
            "sales_21_amount",
            "sales_22_tax",
            "sales_23_zero",
            "sales_24_exempt",
            "special_52_amount",
            "special_53_tax",
            "special_54_amount",
            "special_55_tax",
            "special_84_amount",
            "special_85_tax",
            "special_56_amount",
            "special_57_tax",
            "special_60_amount",
            "special_61_tax",
            "special_62_amount",
            "special_63_amount",
            "special_64_tax",
            "special_65_amount",
            "special_66_tax",
            "sales_total_403",
            "sales_land_amount_26",
            "purchase_28_amount_goods",
            "purchase_29_tax_goods",
            "purchase_30_amount_asset",
            "purchase_31_tax_asset",
            "purchase_32_amount_goods",
            "purchase_33_tax_goods",
            "purchase_34_amount_asset",
            "purchase_35_tax_asset",
            "purchase_36_amount_goods",
            "purchase_37_tax_goods",
            "purchase_38_amount_asset",
            "purchase_39_tax_asset",
            "purchase_78_amount_goods",
            "purchase_79_tax_goods",
            "purchase_80_amount_asset",
            "purchase_81_tax_asset",
            "purchase_40_amount_goods",
            "purchase_41_tax_goods",
            "purchase_42_amount_asset",
            "purchase_43_tax_asset",
            "purchase_44_amount_goods",
            "purchase_45_tax_goods",
            "purchase_46_amount_asset",
            "purchase_47_tax_asset",
            "purchase_48_total_goods",
            "purchase_49_total_asset",
            "nondeductible_ratio_50",
            "input_tax_deductible_51",
            "import_tax_exempt_goods_73",
            "foreign_services_amount_74",
            "foreign_services_tax_75",
            "foreign_services_payable_76",
            "item_101_output_tax_total",
            "item_103_foreign_services",
            "item_104_special_tax",
            "item_105_adjustment_due",
            "item_106_subtotal",
            "item_107_input_tax_total",
            "item_108_previous_credit",
            "item_109_adjustment_refund",
            "item_110_input_tax_subtotal",
            "item_111_net_tax_payable",
            "item_112_net_tax_credit",
            "item_113_refund_limit",
            "item_114_refund_amount",
            "item_115_accumulated_credit",
        )
        return {key: 0 for key in keys}

    def _get_403_special_rate_bucket(self):
        rate = round(float(self.special_tax_rate_percent or 0.0), 2)
        return {
            25.0: ("special_52_amount", "special_53_tax"),
            15.0: ("special_54_amount", "special_55_tax"),
            5.0: ("special_84_amount", "special_85_tax"),
            2.0: ("special_56_amount", "special_57_tax"),
            1.0: ("special_60_amount", "special_61_tax"),
        }.get(rate)

    def _apply_403_sale_entry(self, detail, entry):
        format_code = entry.get("format_code") or ""
        tax_type = entry.get("tax_type") or ""
        amount_untaxed = abs(int(entry.get("amount_untaxed", 0) or 0))
        tax_amount = abs(int(entry.get("tax_amount", 0) or 0))
        if format_code == "31":
            if tax_type == "1":
                detail["sales_1_amount"] += amount_untaxed
                detail["sales_2_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_3_zero"] += amount_untaxed
            elif tax_type == "3":
                detail["sales_4_exempt"] += amount_untaxed
        elif format_code == "35":
            if tax_type == "1":
                detail["sales_5_amount"] += amount_untaxed
                detail["sales_6_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_7_zero"] += amount_untaxed
            elif tax_type == "3":
                detail["sales_8_exempt"] += amount_untaxed
        elif format_code == "32":
            if tax_type == "1":
                detail["sales_9_amount"] += amount_untaxed
                detail["sales_10_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_11_zero"] += amount_untaxed
            elif tax_type == "3":
                detail["sales_12_exempt"] += amount_untaxed
        elif format_code == "36":
            if tax_type == "1":
                detail["sales_13_amount"] += amount_untaxed
                detail["sales_14_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_15_zero"] += amount_untaxed
            elif tax_type == "3":
                detail["sales_16_exempt"] += amount_untaxed
        elif format_code in ("33", "34"):
            if tax_type == "1":
                detail["sales_17_amount"] += amount_untaxed
                detail["sales_18_tax"] += tax_amount
            elif tax_type == "2":
                detail["sales_19_zero"] += amount_untaxed
            elif tax_type == "3":
                detail["sales_20_exempt"] += amount_untaxed
        elif format_code == "37":
            if tax_type == "3":
                detail["special_62_amount"] += amount_untaxed
            else:
                bucket = self._get_403_special_rate_bucket()
                if not bucket:
                    raise UserError(_("403申報書遇到特種稅額資料時，特種稅額稅率只支援1%、2%、5%、15%或25%。"))
                amount_key, tax_key = bucket
                detail[amount_key] += amount_untaxed
                detail[tax_key] += tax_amount
        elif format_code == "38":
            detail["special_63_amount"] += amount_untaxed
            detail["special_64_tax"] += tax_amount

    def _finalize_403_paper_totals(self, detail, rate):
        detail["sales_21_amount"] = (
            detail["sales_1_amount"]
            + detail["sales_5_amount"]
            + detail["sales_9_amount"]
            + detail["sales_13_amount"]
            - detail["sales_17_amount"]
        )
        detail["sales_22_tax"] = (
            detail["sales_2_tax"]
            + detail["sales_6_tax"]
            + detail["sales_10_tax"]
            + detail["sales_14_tax"]
            - detail["sales_18_tax"]
        )
        detail["sales_23_zero"] = (
            detail["sales_3_zero"]
            + detail["sales_7_zero"]
            + detail["sales_11_zero"]
            + detail["sales_15_zero"]
            - detail["sales_19_zero"]
        )
        detail["sales_24_exempt"] = (
            detail["sales_4_exempt"]
            + detail["sales_8_exempt"]
            + detail["sales_12_exempt"]
            + detail["sales_16_exempt"]
            - detail["sales_20_exempt"]
        )
        detail["special_65_amount"] = (
            detail["special_52_amount"]
            + detail["special_54_amount"]
            + detail["special_56_amount"]
            + detail["special_60_amount"]
            + detail["special_84_amount"]
            + detail["special_62_amount"]
            - detail["special_63_amount"]
        )
        detail["special_66_tax"] = (
            detail["special_53_tax"]
            + detail["special_55_tax"]
            + detail["special_57_tax"]
            + detail["special_61_tax"]
            + detail["special_85_tax"]
            - detail["special_64_tax"]
        )
        detail["sales_total_403"] = (
            detail["sales_21_amount"]
            + detail["sales_23_zero"]
            + detail["sales_24_exempt"]
            + detail["special_65_amount"]
        )
        detail["purchase_44_amount_goods"] = (
            detail["purchase_28_amount_goods"]
            + detail["purchase_32_amount_goods"]
            + detail["purchase_36_amount_goods"]
            + detail["purchase_78_amount_goods"]
            - detail["purchase_40_amount_goods"]
        )
        detail["purchase_45_tax_goods"] = (
            detail["purchase_29_tax_goods"]
            + detail["purchase_33_tax_goods"]
            + detail["purchase_37_tax_goods"]
            + detail["purchase_79_tax_goods"]
            - detail["purchase_41_tax_goods"]
        )
        detail["purchase_46_amount_asset"] = (
            detail["purchase_30_amount_asset"]
            + detail["purchase_34_amount_asset"]
            + detail["purchase_38_amount_asset"]
            + detail["purchase_80_amount_asset"]
            - detail["purchase_42_amount_asset"]
        )
        detail["purchase_47_tax_asset"] = (
            detail["purchase_31_tax_asset"]
            + detail["purchase_35_tax_asset"]
            + detail["purchase_39_tax_asset"]
            + detail["purchase_81_tax_asset"]
            - detail["purchase_43_tax_asset"]
        )
        sales_land_amount = max(int(detail.get("sales_land_amount_26", 0) or 0), 0)
        denominator = detail["sales_total_403"] - sales_land_amount
        numerator = detail["sales_24_exempt"] + detail["special_65_amount"] - sales_land_amount
        if denominator > 0 and numerator > 0:
            detail["nondeductible_ratio_50"] = min(int((numerator * 100) / denominator), 100)
        else:
            detail["nondeductible_ratio_50"] = 0
        deductible_base = detail["purchase_45_tax_goods"] + detail["purchase_47_tax_asset"]
        detail["input_tax_deductible_51"] = int((deductible_base * (100 - detail["nondeductible_ratio_50"])) / 100)
        detail["item_101_output_tax_total"] = detail["sales_22_tax"]
        detail["item_103_foreign_services"] = detail["foreign_services_payable_76"]
        detail["item_104_special_tax"] = detail["special_66_tax"]
        detail["item_105_adjustment_due"] = 0
        detail["item_106_subtotal"] = (
            detail["item_101_output_tax_total"]
            + detail["item_103_foreign_services"]
            + detail["item_104_special_tax"]
            + detail["item_105_adjustment_due"]
        )
        detail["item_107_input_tax_total"] = detail["input_tax_deductible_51"]
        detail["item_108_previous_credit"] = 0
        detail["item_109_adjustment_refund"] = 0
        detail["item_110_input_tax_subtotal"] = (
            detail["item_107_input_tax_total"]
            + detail["item_108_previous_credit"]
            + detail["item_109_adjustment_refund"]
        )
        detail["item_111_net_tax_payable"] = max(
            detail["item_106_subtotal"] - detail["item_110_input_tax_subtotal"],
            0,
        )
        detail["item_112_net_tax_credit"] = max(
            detail["item_110_input_tax_subtotal"] - detail["item_106_subtotal"],
            0,
        )
        detail["item_113_refund_limit"] = max(int(round(detail["sales_23_zero"] * rate)), 0) + max(
            detail["purchase_47_tax_asset"],
            0,
        )
        detail["item_114_refund_amount"] = min(
            detail["item_112_net_tax_credit"],
            detail["item_113_refund_limit"],
        )
        detail["item_115_accumulated_credit"] = max(
            detail["item_112_net_tax_credit"] - detail["item_114_refund_amount"],
            0,
        )

    def _prepare_paper_run(self):
        company = self.company_id
        if not company.vat:
            raise UserError(_("請先在公司資料填寫統一編號(8碼)"))
        if not company.tw_tax_id_9 or len(_digits_only(company.tw_tax_id_9)) != 9:
            raise UserError(_("請先在公司資料填寫稅籍編號(9碼)"))

        date_from, date_to = self._period_range()
        moves = self._get_moves(date_from, date_to)
        exported = []
        rate = float(self.tax_rate_percent or 0.0) / 100.0
        taxable_sales = 0
        zero_rate_sales = 0
        exempt_sales = 0
        output_tax = 0
        input_tax_deductible = 0
        input_tax_deductible_goods = 0
        input_tax_deductible_asset = 0
        sales_invoice_count = 0
        form_401_totals = self._init_401_paper_totals() if self.paper_form_type == "401" else False
        form_403_totals = self._init_403_paper_totals() if self.paper_form_type == "403" else False

        for move in moves.sorted(lambda record: (record.invoice_date or record.date, record.name)):
            if not move.tw_blr_format_code:
                exported.append((move.display_name, "跳過：缺申報格式代號"))
                continue
            entries = self._build_export_entries(move)
            if not entries:
                exported.append((move.display_name, "跳過：沒有可統計的有效資料"))
                continue
            if move.move_type == "out_invoice":
                sales_invoice_count += 1
            for entry in entries:
                amount_untaxed = int(entry.get("amount_untaxed", 0) or 0)
                tax_amount = int(entry.get("tax_amount", 0) or 0)
                tax_type = entry.get("tax_type") or ""
                deduct_code = entry.get("deduct_code") or " "
                if move.move_type in ("out_invoice", "out_refund"):
                    if tax_type == "1":
                        taxable_sales += amount_untaxed
                        output_tax += tax_amount
                    elif tax_type == "2":
                        zero_rate_sales += amount_untaxed
                    elif tax_type == "3":
                        exempt_sales += amount_untaxed
                    if form_401_totals:
                        self._apply_401_sale_entry(form_401_totals, entry)
                    if form_403_totals:
                        self._apply_403_sale_entry(form_403_totals, entry)
                else:
                    if deduct_code in ("1", "2"):
                        input_tax_deductible += tax_amount
                        if deduct_code == "1":
                            input_tax_deductible_goods += tax_amount
                        else:
                            input_tax_deductible_asset += tax_amount
                    if form_401_totals:
                        self._apply_401_purchase_entry(form_401_totals, entry)
                    if form_403_totals:
                        self._apply_401_purchase_entry(form_403_totals, entry)
            status = "OK" if len(entries) == 1 else _("OK，已依進貨及費用/固定資產拆分%s筆") % len(entries)
            exported.append((move.display_name, status))

        if form_401_totals:
            self._finalize_401_paper_totals(form_401_totals)
            taxable_sales = form_401_totals["sales_21_amount"]
            zero_rate_sales = form_401_totals["sales_23_zero"]
            output_tax = form_401_totals["sales_22_tax"]
            input_tax_deductible_goods = form_401_totals["purchase_45_tax_goods"]
            input_tax_deductible_asset = form_401_totals["purchase_47_tax_asset"]
            input_tax_deductible = input_tax_deductible_goods + input_tax_deductible_asset
        elif form_403_totals:
            self._finalize_403_paper_totals(form_403_totals, rate)
            taxable_sales = form_403_totals["sales_21_amount"]
            zero_rate_sales = form_403_totals["sales_23_zero"]
            exempt_sales = form_403_totals["sales_24_exempt"]
            output_tax = form_403_totals["item_101_output_tax_total"]
            input_tax_deductible_goods = form_403_totals["purchase_45_tax_goods"]
            input_tax_deductible_asset = form_403_totals["purchase_47_tax_asset"]
            input_tax_deductible = form_403_totals["input_tax_deductible_51"]

        input_tax_subtotal = input_tax_deductible
        refund_limit = max(int(round(zero_rate_sales * rate)), 0) + max(input_tax_deductible_asset, 0)
        net_tax_payable = max(output_tax - input_tax_subtotal, 0)
        net_tax_credit = max(input_tax_subtotal - output_tax, 0)
        refund_amount = min(net_tax_credit, refund_limit)
        accumulated_credit = max(net_tax_credit - refund_amount, 0)
        totals = {
            "invoice_count": sales_invoice_count,
            "taxable_sales": taxable_sales,
            "zero_rate_sales": zero_rate_sales,
            "exempt_sales": exempt_sales,
            "sales_total_401": form_401_totals["sales_25_total"] if form_401_totals else taxable_sales + zero_rate_sales,
            "sales_total_403": form_403_totals["sales_total_403"] if form_403_totals else taxable_sales + zero_rate_sales + exempt_sales,
            "output_tax": output_tax,
            "input_tax_deductible": input_tax_deductible,
            "input_tax_deductible_goods": input_tax_deductible_goods,
            "input_tax_deductible_asset": input_tax_deductible_asset,
            "previous_credit": 0,
            "input_tax_subtotal": input_tax_subtotal,
            "refund_limit": refund_limit,
            "refund_amount": refund_amount,
            "net_tax_payable": net_tax_payable,
            "net_tax_credit": net_tax_credit,
            "accumulated_credit": accumulated_credit,
        }
        if form_401_totals:
            totals.update(form_401_totals)
        if form_403_totals:
            totals.update(form_403_totals)
            totals.update({
                "output_tax": form_403_totals["item_101_output_tax_total"],
                "input_tax_deductible": form_403_totals["input_tax_deductible_51"],
                "input_tax_subtotal": form_403_totals["item_110_input_tax_subtotal"],
                "refund_limit": form_403_totals["item_113_refund_limit"],
                "refund_amount": form_403_totals["item_114_refund_amount"],
                "net_tax_payable": form_403_totals["item_111_net_tax_payable"],
                "net_tax_credit": form_403_totals["item_112_net_tax_credit"],
                "accumulated_credit": form_403_totals["item_115_accumulated_credit"],
            })
        return {
            "company": company,
            "moves": moves,
            "errors": [],
            "exported": exported,
            "txt_lines": [],
            "totals": totals,
        }
    def _build_check_report(self, company, exported, totals):
        report = [
            "營業稅申報匯出檢核報表",
            f"統一編號：{_zfill_digits((company.vat or '').replace('TW', ''), 8)}",
            f"稅籍編號(9碼)：{_digits_only(company.tw_tax_id_9)}",
            f"申報期別(民國)：{_zfill_digits(self.year_roc, 3)}{_zfill_digits(self.month, 2)}",
            "",
            f"明細行數(BAN.TXT)：{totals.get('invoice_count', 0)}",
            f"應稅銷售額(未稅)：{totals.get('taxable_sales', 0)}",
            f"零稅率銷售額：{totals.get('zero_rate_sales', 0)}",
            f"免稅銷售額：{totals.get('exempt_sales', 0)}",
            f"銷項稅額：{totals.get('output_tax', 0)}",
            f"進項可扣抵稅額：{totals.get('input_tax_deductible', 0)}",
            "",
            "逐筆結果：",
        ]
        for name, status in exported:
            report.append(f"- {name}：{status}")
        return "\r\n".join(report) + "\r\n"
    def _get_builtin_paper_xlsx_template_spec(self):
        xlsx_path = Path(__file__).resolve().parents[1] / "static" / "templates" / "official_forms" / "vat_forms_official.xlsx"
        if not xlsx_path.is_file():
            raise UserError(_("模組內建官方空白申報書XLSX不存在，請確認 static/templates/official_forms/vat_forms_official.xlsx"))
        with xlsx_path.open("rb") as xlsx_file:
            return {
                "template": xlsx_file.read(),
                "template_name": "vat_forms_official.xlsx",
                "template_source": "module_builtin",
            }

    def _get_company_address_text(self):
        company = self.company_id
        address_parts = [
            company.zip,
            company.state_id.name,
            company.city,
            company.street,
            company.street2,
        ]
        return _clean("".join(part for part in address_parts if part))

    def _get_filer_phone_text(self):
        company = self.company_id
        parts = []
        area = _clean(company.tw_filer_tel_area)
        tel = _clean(company.tw_filer_tel)
        ext = _clean(company.tw_filer_tel_ext)
        if area:
            parts.append(area)
        if tel:
            parts.append(tel)
        phone = "-".join(parts)
        if ext:
            phone = f"{phone}#{ext}" if phone else ext
        return phone

    def _get_responsible_name_text(self):
        return _clean(self.company_id.tw_responsible_name)

    def _validate_paper_export_fields(self):
        missing_fields = []
        if not _clean(self.company_id.name):
            missing_fields.append("營業人名稱")
        if len(_digits_only((self.company_id.vat or "").replace("TW", ""))) != 8:
            missing_fields.append("統一編號(8碼)")
        if len(_digits_only(self.company_id.tw_tax_id_9)) != 9:
            missing_fields.append("稅籍編號(9碼)")
        if not self._get_responsible_name_text():
            missing_fields.append("負責人姓名")
        if not self._get_company_address_text():
            missing_fields.append("營業地址")
        if not _clean(self.company_id.tw_filer_name):
            missing_fields.append("申報人姓名")
        if not _clean(self.company_id.tw_filer_tel):
            missing_fields.append("申報人電話")
        if missing_fields:
            raise UserError(
                _("產生紙本申報表前，請先補齊下列欄位：\n- %s") % ("\n- ".join(missing_fields))
            )
        return "委任申報" if _clean(self.company_id.tw_agent_reg_no) else "自行申報"

    def _get_paper_payload(self, totals):
        filing_date = date.today()
        filing_date_text = f"{filing_date.year - 1911}/{filing_date.month:02d}/{filing_date.day:02d}"
        filer_name = _clean(self.company_id.tw_filer_name)
        filer_idno = _clean(self.company_id.tw_filer_idno)
        filer_phone = self._get_filer_phone_text()
        agent_reg_no = _clean(self.company_id.tw_agent_reg_no)
        use_agent_row = bool(agent_reg_no)
        filing_mode = "委任申報" if use_agent_row else "自行申報"
        payload = {
            "vat_no": _zfill_digits((self.company_id.vat or "").replace("TW", ""), 8),
            "company_name": _clean(self.company_id.name),
            "tax_id_9": _zfill_digits(self.company_id.tw_tax_id_9, 9),
            "roc_year": _zfill_digits(self.year_roc, 3),
            "period_month": _format_period_month(self.month),
            "responsible_name": self._get_responsible_name_text(),
            "company_address": self._get_company_address_text(),
            "invoice_count": str(int(totals.get("invoice_count", 0))),
            "taxable_sales": str(int(totals.get("taxable_sales", 0))),
            "zero_rate_sales": str(int(totals.get("zero_rate_sales", 0))),
            "exempt_sales": str(int(totals.get("exempt_sales", 0))),
            "sales_total": str(
                int(
                    totals.get("sales_total_403", 0)
                    if self.paper_form_type == "403"
                    else totals.get("sales_total_401", 0)
                )
            ),
            "output_tax": str(int(totals.get("output_tax", 0))),
            "input_tax_goods": str(int(totals.get("input_tax_deductible_goods", 0))),
            "input_tax_asset": str(int(totals.get("input_tax_deductible_asset", 0))),
            "input_tax_total": str(int(totals.get("input_tax_deductible", 0))),
            "previous_credit": str(int(totals.get("previous_credit", 0))),
            "input_tax_subtotal": str(int(totals.get("input_tax_subtotal", 0))),
            "net_tax_payable": str(int(totals.get("net_tax_payable", 0))),
            "net_tax_credit": str(int(totals.get("net_tax_credit", 0))),
            "refund_limit": str(int(totals.get("refund_limit", 0))),
            "refund_amount": str(int(totals.get("refund_amount", 0))),
            "accumulated_credit": str(int(totals.get("accumulated_credit", 0))),
            "filing_date": filing_date_text,
            "filing_mode": filing_mode,
            "self_filer_name": "" if use_agent_row else filer_name,
            "self_filer_idno": "" if use_agent_row else filer_idno,
            "self_filer_phone": "" if use_agent_row else filer_phone,
            "self_filer_reg_no": "",
            "agent_name": filer_name if use_agent_row else "",
            "agent_idno": filer_idno if use_agent_row else "",
            "agent_phone": filer_phone if use_agent_row else "",
            "agent_reg_no": agent_reg_no if use_agent_row else "",
        }
        if self.paper_form_type == "401":
            detail_keys = (
                "sales_21_amount",
                "sales_22_tax",
                "sales_23_zero",
                "sales_25_total",
                "sales_27_fixed_asset",
                "purchase_28_amount_goods",
                "purchase_29_tax_goods",
                "purchase_30_amount_asset",
                "purchase_31_tax_asset",
                "purchase_32_amount_goods",
                "purchase_33_tax_goods",
                "purchase_34_amount_asset",
                "purchase_35_tax_asset",
                "purchase_36_amount_goods",
                "purchase_37_tax_goods",
                "purchase_38_amount_asset",
                "purchase_39_tax_asset",
                "purchase_40_amount_goods",
                "purchase_41_tax_goods",
                "purchase_42_amount_asset",
                "purchase_43_tax_asset",
                "purchase_44_amount_goods",
                "purchase_45_tax_goods",
                "purchase_46_amount_asset",
                "purchase_47_tax_asset",
                "purchase_48_total_goods",
                "purchase_49_total_asset",
                "purchase_78_amount_goods",
                "purchase_79_tax_goods",
                "purchase_80_amount_asset",
                "purchase_81_tax_asset",
                "import_tax_exempt_goods",
                "purchase_foreign_services",
            )
            for key in detail_keys:
                payload[key] = str(int(totals.get(key, 0)))
            payload.update({
                "item_101_output_tax": str(int(totals.get("output_tax", 0))),
                "item_107_input_tax_total": str(int(totals.get("input_tax_deductible", 0))),
                "item_108_previous_credit": str(int(totals.get("previous_credit", 0))),
                "item_110_input_tax_subtotal": str(int(totals.get("input_tax_subtotal", 0))),
                "item_111_net_tax_payable": str(int(totals.get("net_tax_payable", 0))),
                "item_112_net_tax_credit": str(int(totals.get("net_tax_credit", 0))),
                "item_113_refund_limit": str(int(totals.get("refund_limit", 0))),
                "item_114_refund_amount": str(int(totals.get("refund_amount", 0))),
                "item_115_accumulated_credit": str(int(totals.get("accumulated_credit", 0))),
            })
        elif self.paper_form_type == "403":
            detail_keys = (
                "sales_1_amount",
                "sales_2_tax",
                "sales_3_zero",
                "sales_4_exempt",
                "sales_5_amount",
                "sales_6_tax",
                "sales_7_zero",
                "sales_8_exempt",
                "sales_9_amount",
                "sales_10_tax",
                "sales_11_zero",
                "sales_12_exempt",
                "sales_13_amount",
                "sales_14_tax",
                "sales_15_zero",
                "sales_16_exempt",
                "sales_17_amount",
                "sales_18_tax",
                "sales_19_zero",
                "sales_20_exempt",
                "sales_21_amount",
                "sales_22_tax",
                "sales_23_zero",
                "sales_24_exempt",
                "special_52_amount",
                "special_53_tax",
                "special_54_amount",
                "special_55_tax",
                "special_84_amount",
                "special_85_tax",
                "special_56_amount",
                "special_57_tax",
                "special_60_amount",
                "special_61_tax",
                "special_62_amount",
                "special_63_amount",
                "special_64_tax",
                "special_65_amount",
                "special_66_tax",
                "sales_total_403",
                "sales_land_amount_26",
                "purchase_28_amount_goods",
                "purchase_29_tax_goods",
                "purchase_30_amount_asset",
                "purchase_31_tax_asset",
                "purchase_32_amount_goods",
                "purchase_33_tax_goods",
                "purchase_34_amount_asset",
                "purchase_35_tax_asset",
                "purchase_36_amount_goods",
                "purchase_37_tax_goods",
                "purchase_38_amount_asset",
                "purchase_39_tax_asset",
                "purchase_78_amount_goods",
                "purchase_79_tax_goods",
                "purchase_80_amount_asset",
                "purchase_81_tax_asset",
                "purchase_40_amount_goods",
                "purchase_41_tax_goods",
                "purchase_42_amount_asset",
                "purchase_43_tax_asset",
                "purchase_44_amount_goods",
                "purchase_45_tax_goods",
                "purchase_46_amount_asset",
                "purchase_47_tax_asset",
                "purchase_48_total_goods",
                "purchase_49_total_asset",
                "nondeductible_ratio_50",
                "input_tax_deductible_51",
                "import_tax_exempt_goods_73",
                "foreign_services_amount_74",
                "foreign_services_tax_75",
                "foreign_services_payable_76",
                "item_101_output_tax_total",
                "item_103_foreign_services",
                "item_104_special_tax",
                "item_105_adjustment_due",
                "item_106_subtotal",
                "item_107_input_tax_total",
                "item_108_previous_credit",
                "item_109_adjustment_refund",
                "item_110_input_tax_subtotal",
                "item_111_net_tax_payable",
                "item_112_net_tax_credit",
                "item_113_refund_limit",
                "item_114_refund_amount",
                "item_115_accumulated_credit",
            )
            for key in detail_keys:
                payload[key] = str(int(totals.get(key, 0)))
        return payload
    def _xlsx_set_inline_string(self, cell, value):
        for child in list(cell):
            cell.remove(child)
        cell.set("t", "inlineStr")
        inline_string = ET.SubElement(cell, _xlsx_tag("is"))
        text = ET.SubElement(inline_string, _xlsx_tag("t"))
        text.text = str(value or "")
        if text.text != text.text.strip() or "\n" in text.text or text.text == "":
            text.set(XML_SPACE_ATTR, "preserve")

    def _xlsx_apply_cell_value(self, cells, cell_ref, value, missing_cells, written_fields):
        cell = cells.get(cell_ref)
        if cell is None:
            missing_cells.append(cell_ref)
            return
        self._xlsx_set_inline_string(cell, value)
        written_fields.append(cell_ref)

    def _xlsx_apply_digit_field(self, cells, cell_refs, value, missing_cells, written_fields):
        chars = list(str(value or ""))
        if len(chars) < len(cell_refs):
            chars.extend([""] * (len(cell_refs) - len(chars)))
        for cell_ref, char in zip(cell_refs, chars[: len(cell_refs)]):
            self._xlsx_apply_cell_value(cells, cell_ref, char, missing_cells, written_fields)

    def _paper_field_default_value(self, field_name, value):
        if value not in (None, ""):
            return value
        if field_name in {
            "company_name",
            "responsible_name",
            "company_address",
            "filing_date",
            "self_filer_name",
            "self_filer_idno",
            "self_filer_phone",
            "self_filer_reg_no",
            "agent_name",
            "agent_idno",
            "agent_phone",
            "agent_reg_no",
        }:
            return ""
        return "0"

    def _build_paper_xlsx(self, template_bytes, payload):
        config = PAPER_XLSX_CONFIG[self.paper_form_type]
        missing_cells = []
        written_fields = []
        input_buffer = io.BytesIO(template_bytes)
        output_buffer = io.BytesIO()

        with zipfile.ZipFile(input_buffer, "r") as source_zip:
            with zipfile.ZipFile(output_buffer, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
                for info in source_zip.infolist():
                    data = source_zip.read(info.filename)
                    if info.filename == config["sheet_path"]:
                        sheet_root = ET.fromstring(data)
                        cells = {cell.get("r"): cell for cell in sheet_root.findall(f".//{_xlsx_tag('c')}")}
                        for field_name, cell_refs in config.get("digit_fields", {}).items():
                            self._xlsx_apply_digit_field(
                                cells,
                                cell_refs,
                                payload.get(field_name, ""),
                                missing_cells,
                                written_fields,
                            )
                        for field_name, cell_ref in config.get("text_fields", {}).items():
                            value = self._paper_field_default_value(field_name, payload.get(field_name))
                            if value in (None, ""):
                                continue
                            self._xlsx_apply_cell_value(cells, cell_ref, value, missing_cells, written_fields)
                        data = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
                    target_zip.writestr(info, data)

        return output_buffer.getvalue(), {
            "sheet_name": config["sheet_name"],
            "sheet_path": config["sheet_path"],
            "written_fields": written_fields,
            "missing_cells": missing_cells,
            "note": config.get("note", ""),
        }

    def _build_paper_xlsx_report(self, template_name, payload, xlsx_meta):
        report_lines = [
            "營業稅申報書 Excel 匯出結果",
            f"申報書類型：{self.paper_form_type}",
            f"模板檔名：{template_name}",
            f"工作表：{xlsx_meta['sheet_name']}",
            "模板來源：模組內建官方空白申報書(XLSX)",
            f"申報方式：{payload.get('filing_mode', '')}",
            f"申報期別(民國)：{_zfill_digits(self.year_roc, 3)}{_zfill_digits(self.month, 2)}",
            f"已寫入儲存格數：{len(xlsx_meta['written_fields'])}",
            f"應稅銷售額：{payload.get('taxable_sales', '')}",
            f"零稅率銷售額：{payload.get('zero_rate_sales', '')}",
            f"免稅銷售額：{payload.get('exempt_sales', '')}",
            f"銷項稅額：{payload.get('output_tax', '')}",
            f"可扣抵進項稅額：{payload.get('input_tax_total', '')}",
            f"本期應實繳稅額：{payload.get('net_tax_payable', '')}",
            f"本期留抵稅額：{payload.get('net_tax_credit', '')}",
        ]
        if xlsx_meta["note"]:
            report_lines.extend(["", xlsx_meta["note"]])
        if xlsx_meta["missing_cells"]:
            report_lines.extend([
                "",
                "下列儲存格在官方模板中未找到，需檢查政府範本是否改版：",
                ", ".join(xlsx_meta["missing_cells"]),
            ])
        report_lines.extend([
            "",
            "XLSX已直接寫入官方空白申報書工作表，建議先以Excel或LibreOffice開啟後檢查列印結果。",
        ])
        return "\r\n".join(report_lines)

    def action_generate_paper_xlsx(self):
        self._validate_paper_export_fields()

        template_spec = self._get_builtin_paper_xlsx_template_spec()
        original_form_type = self.paper_form_type or "401"
        xlsx_bytes = template_spec["template"]
        form_payloads = {}
        form_meta = {}

        try:
            for form_type in ("401", "403", "404"):
                self.paper_form_type = form_type
                export_data = self._prepare_paper_run()
                payload = self._get_paper_payload(export_data["totals"])
                xlsx_bytes, meta = self._build_paper_xlsx(xlsx_bytes, payload)
                form_payloads[form_type] = payload
                form_meta[form_type] = meta
        finally:
            self.paper_form_type = original_form_type

        self.paper_xlsx = base64.b64encode(xlsx_bytes)
        self.paper_xlsx_name = (
            f"{self.company_id.id}_VAT_{_zfill_digits(self.year_roc, 3)}"
            f"{_zfill_digits(self.month, 2)}.xlsx"
        )

        report_lines = [
            "進銷項核對表 Excel 匯出結果",
            f"模板檔名：{template_spec['template_name']}",
            f"申報期別(民國)：{_zfill_digits(self.year_roc, 3)}{_zfill_digits(self.month, 2)}",
            "模板來源：模組內建官方空白申報書(XLSX)",
        ]
        for form_type in ("401", "403", "404"):
            payload = form_payloads.get(form_type, {})
            meta = form_meta.get(form_type, {})
            report_lines.extend([
                "",
                f"[{form_type}] 工作表：{meta.get('sheet_name', '')}",
                f"已寫入儲存格數：{len(meta.get('written_fields', []))}",
                f"應稅銷售額：{payload.get('taxable_sales', '')}",
                f"零稅率銷售額：{payload.get('zero_rate_sales', '')}",
                f"免稅銷售額：{payload.get('exempt_sales', '')}",
                f"銷項稅額：{payload.get('output_tax', '')}",
                f"可扣抵進項稅額：{payload.get('input_tax_total', '')}",
                f"本期應實繳稅額：{payload.get('net_tax_payable', '')}",
                f"本期留抵稅額：{payload.get('net_tax_credit', '')}",
            ])
            if meta.get('note'):
                report_lines.extend(["", meta['note']])
            if meta.get('missing_cells'):
                report_lines.extend([
                    "",
                    "下列儲存格在官方模板中未找到，需檢查政府範本是否改版：",
                    ", ".join(meta['missing_cells']),
                ])
        report_lines.extend([
            "",
            "XLSX已直接寫入官方空白申報書的401、403、404工作表，建議先以Excel或LibreOffice開啟後檢查列印結果。",
        ])
        self.paper_check_report = "\r\n".join(report_lines)
        return {
            "type": "ir.actions.act_window",
            "res_model": "tw.vat.filing.wizard",
            "view_mode": "form",
            "res_id": self.id,
            "target": "new",
        }

    def action_generate_zip(self):
        export_data = self._prepare_export_run()
        if export_data["errors"]:
            raise UserError(_("申報匯出失敗，請先修正：\n- %s") % ("\n- ".join(export_data["errors"])))

        company = export_data["company"]
        vat_8 = _zfill_digits((company.vat or "").replace("TW", ""), 8)
        txt_name = f"{vat_8}.TXT"
        tet_name = f"{vat_8}.TET_U"
        txt_content = ("\r\n".join(export_data["txt_lines"]) + ("\r\n" if export_data["txt_lines"] else "")).encode("utf-8")
        tet_content = (self._tet_u_line(export_data["totals"]) + "\r\n").encode("utf-8")
        report_text = self._build_check_report(company, export_data["exported"], export_data["totals"])
        if not export_data["moves"]:
            report_text = report_text.rstrip("\r\n") + "\r\n- 本期無已過帳單據\r\n"
        self.check_report = report_text

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr(txt_name, txt_content)
            zip_file.writestr(tet_name, tet_content)
            zip_file.writestr("CHECK_REPORT.txt", report_text.encode("utf-8"))

        self.export_zip = base64.b64encode(buffer.getvalue())
        self.export_zip_name = f"{vat_8}_VAT_FILING_{_zfill_digits(self.year_roc, 3)}{_zfill_digits(self.month, 2)}.zip"
        return {
            "type": "ir.actions.act_window",
            "res_model": "tw.vat.filing.wizard",
            "view_mode": "form",
            "res_id": self.id,
            "target": "new",
        }








