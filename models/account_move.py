import re

from odoo import _, api, fields, models
from odoo.exceptions import ValidationError


IN_FORMAT_SELECTION = [
    ("21", "21 進項三聯式、電子計算機統一發票"),
    ("22", "22 進項二聯式收銀機統一發票、載有稅額之其他憑證"),
    (
        "23",
        "23 進貨退出或折讓證明單(三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票)",
    ),
    ("24", "24 進貨退出或折讓證明單(二聯式收銀機統一發票及載有稅額之其他憑證)"),
    ("25", "25 進項三聯式收銀機統一發票及一般稅額計算之電子發票(含公用事業載具流水號)"),
    ("26", "26 彙總登錄：每張稅額500元以下之進項三聯式、電子計算機統一發票"),
    ("27", "27 彙總登錄：每張稅額500元以下之進項二聯式收銀機統一發票、載有稅額之其他憑證"),
    ("28", "28 進項海關代徵營業稅繳納證"),
    ("29", "29 進項海關退還溢繳營業稅申報單"),
]

OUT_FORMAT_SELECTION = [
    ("31", "31 銷項三聯式、電子計算機統一發票"),
    ("32", "32 銷項二聯式、二聯式收銀機統一發票"),
    (
        "33",
        "33 銷貨退回或折讓證明單(三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票)",
    ),
    ("34", "34 銷貨退回或折讓證明單(二聯式、二聯式收銀機統一發票及銷項免用統一發票)"),
    ("35", "35 銷項三聯式收銀機統一發票及一般稅額計算之電子發票"),
    ("36", "36 銷項免用統一發票"),
]

ALL_FORMAT_SELECTION = IN_FORMAT_SELECTION + OUT_FORMAT_SELECTION + [
    ("37", "37 銷項憑證、特種稅額計算之電子發票(特種稅額計算)"),
    ("38", "38 銷貨退回或折讓證明單(特種稅額計算)"),
]

INV_OR_UTILITY_FORMATS = {"25", "35"}
INV_OR_OTHER_FORMATS = {"22", "24", "27", "32", "34", "36"}
INV_ONLY_FORMATS = {"21", "23", "26", "31", "33"}
CUSTOMS_FORMATS = {"28", "29"}
UNSUPPORTED_FORMATS = {"37", "38"}
IN_FORMAT_CODES = {code for code, _ in IN_FORMAT_SELECTION}
OUT_FORMAT_CODES = {code for code, _ in OUT_FORMAT_SELECTION}
ZERO_RATE_KEYWORDS = ("零稅率", "zero rate", "zero-rate", "zero rated", "zerorated")
EXEMPT_KEYWORDS = ("免稅", "免徵", "exempt")
NON_DEDUCTIBLE_KEYWORDS = (
    "不可扣抵",
    "不得扣抵",
    "不可抵扣",
    "不得抵扣",
    "不予扣抵",
    "不予抵扣",
    "non-deductible",
    "non deductible",
    "nondeductible",
)
DEDUCTIBLE_KEYWORDS = ("可扣抵", "可抵扣", "準予扣抵", "準予抵扣", "deductible")
CLUE_TRANSLATION_MAP = str.maketrans({
    "\u7a0e": "稅",
    "\u5f81": "徵",
    "\u51c6": "準",
    "\u53f0": "臺",
})


class AccountMove(models.Model):
    _inherit = "account.move"

    tw_blr_auto = fields.Boolean(string="申報欄位自動帶入", default=True)
    tw_blr_skip_export = fields.Boolean(string="不匯出申報資料", default=False)
    tw_blr_invoice_autofill_trigger = fields.Boolean(
        string="自動帶入發票字軌與號碼",
        compute="_compute_tw_blr_invoice_autofill_trigger",
        inverse="_inverse_tw_blr_invoice_autofill_trigger",
        store=False,
    )
    tw_blr_tax_autofill_trigger = fields.Boolean(
        string="自動帶入課稅別與扣抵",
        compute="_compute_tw_blr_tax_autofill_trigger",
        inverse="_inverse_tw_blr_tax_autofill_trigger",
        store=False,
    )
    tw_blr_format_code = fields.Selection(selection=ALL_FORMAT_SELECTION, string="申報格式代號")
    tw_blr_format_code_in = fields.Selection(
        selection=IN_FORMAT_SELECTION,
        string="申報格式代號",
        compute="_compute_tw_blr_format_code_in",
        inverse="_inverse_tw_blr_format_code_in",
        store=False,
    )
    tw_blr_format_code_out = fields.Selection(
        selection=OUT_FORMAT_SELECTION,
        string="申報格式代號",
        compute="_compute_tw_blr_format_code_out",
        inverse="_inverse_tw_blr_format_code_out",
        store=False,
    )
    tw_invoice_track = fields.Char(string="發票字軌(2碼)")
    tw_invoice_number = fields.Char(string="發票號碼(8碼)")
    tw_other_voucher_no = fields.Char(string="其他憑證號碼(10碼)")
    tw_utility_carrier_no = fields.Char(string="公用事業載具流水號(BB+8碼)")
    tw_customs_pay_no = fields.Char(string="海關代徵營業稅繳納證號碼(14碼)")
    tw_tax_type = fields.Selection(
        selection=[("1", "1 應稅"), ("2", "2 零稅率"), ("3", "3 免稅")],
        string="課稅別",
        default="1",
    )
    tw_deduct_code = fields.Selection(
        selection=[
            ("1", "1 進項可扣抵之進貨及費用"),
            ("2", "2 進項可扣抵之固定資產"),
            ("3", "3 進項不可扣抵之進貨及費用"),
            ("4", "4 進項不可扣抵之固定資產"),
        ],
        string="抵扣代號",
        default="1",
    )

    @api.depends("tw_blr_format_code")
    def _compute_tw_blr_format_code_in(self):
        for move in self:
            move.tw_blr_format_code_in = move.tw_blr_format_code if move.tw_blr_format_code in IN_FORMAT_CODES else False

    def _inverse_tw_blr_format_code_in(self):
        for move in self:
            move.tw_blr_format_code = move.tw_blr_format_code_in or False

    @api.depends("tw_blr_format_code")
    def _compute_tw_blr_format_code_out(self):
        for move in self:
            move.tw_blr_format_code_out = move.tw_blr_format_code if move.tw_blr_format_code in OUT_FORMAT_CODES else False

    def _inverse_tw_blr_format_code_out(self):
        for move in self:
            move.tw_blr_format_code = move.tw_blr_format_code_out or False

    def _compute_tw_blr_invoice_autofill_trigger(self):
        for move in self:
            move.tw_blr_invoice_autofill_trigger = False

    def _inverse_tw_blr_invoice_autofill_trigger(self):
        return

    def _compute_tw_blr_tax_autofill_trigger(self):
        for move in self:
            move.tw_blr_tax_autofill_trigger = False

    def _inverse_tw_blr_tax_autofill_trigger(self):
        return

    def _blr_build_onchange_warning(self, messages):
        messages = list(dict.fromkeys(message for message in messages if message))
        if not messages:
            return {}
        return {
            "warning": {
                "title": _("自動帶入提示"),
                "message": "\n".join(messages),
            }
        }

    def _blr_get_invoice_autofill_values(self):
        self.ensure_one()
        return {}

    def _blr_apply_invoice_autofill(self):
        self.ensure_one()
        values = self._blr_get_invoice_autofill_values() or {}
        track = (values.get("tw_invoice_track") or "").strip().upper()
        number = re.sub(r"\D", "", str(values.get("tw_invoice_number") or ""))

        if not track and not number:
            return _("目前沒有可用的整合欄位可帶入發票字軌與號碼。")
        if len(track) != 2 or len(number) != 8:
            return _("整合欄位提供的發票字軌或號碼格式不正確。")

        self.tw_invoice_track = track
        self.tw_invoice_number = number
        self.tw_other_voucher_no = False
        self.tw_utility_carrier_no = False
        return False

    def _blr_tax_amount_is_zero(self):
        self.ensure_one()
        currency = self.currency_id or self.company_currency_id
        if currency:
            return currency.is_zero(self.amount_tax)
        return not self.amount_tax

    def _blr_relevant_invoice_lines(self):
        self.ensure_one()
        lines = self.invoice_line_ids.filtered(lambda line: not line.display_type)
        if lines:
            return lines
        return self.line_ids.filtered(
            lambda line: not line.display_type
            and not line.tax_line_id
            and getattr(line.account_id, "account_type", "") not in ("asset_receivable", "liability_payable")
        )

    def _blr_collect_invoice_taxes(self):
        self.ensure_one()
        taxes = self.env["account.tax"]
        for line in self._blr_relevant_invoice_lines():
            taxes |= line.tax_ids
        taxes |= self.line_ids.filtered(lambda line: line.tax_line_id).mapped("tax_line_id")
        return taxes

    def _blr_collect_tax_tags(self):
        self.ensure_one()
        tags = self.env["account.account.tag"]
        tax_lines = self.line_ids.filtered(lambda line: line.tax_line_id)
        for line in self._blr_relevant_invoice_lines() | tax_lines:
            tags |= line.tax_tag_ids
        for tax in self._blr_collect_invoice_taxes():
            tags |= tax.invoice_repartition_line_ids.mapped("tag_ids")
            tags |= tax.refund_repartition_line_ids.mapped("tag_ids")
        return tags

    def _blr_tax_keyword_text(self, tax):
        parts = []
        for attr in ("name", "description", "invoice_label"):
            value = getattr(tax, attr, False)
            if value:
                parts.append(str(value))
        tax_group = getattr(tax, "tax_group_id", False)
        if tax_group and getattr(tax_group, "name", False):
            parts.append(str(tax_group.name))
        return " ".join(parts).lower()

    def _blr_tax_clues(self):
        self.ensure_one()
        clues = [self._blr_tax_keyword_text(tax) for tax in self._blr_collect_invoice_taxes()]
        clues.extend(str(tag.name).lower() for tag in self._blr_collect_tax_tags() if tag.name)
        if self.fiscal_position_id and self.fiscal_position_id.name:
            clues.append(str(self.fiscal_position_id.name).lower())
        return [clue for clue in clues if clue]

    def _blr_normalize_clue_text(self, text):
        normalized = (text or "").lower().translate(CLUE_TRANSLATION_MAP)
        return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", normalized)

    def _blr_clue_contains_keyword(self, clue, keyword):
        clue_text = (clue or "").lower()
        keyword_text = (keyword or "").strip().lower()
        if not clue_text or not keyword_text:
            return False
        if re.fullmatch(r"[a-z0-9\s\-_]+", keyword_text):
            parts = [re.escape(part) for part in re.split(r"[\s\-_]+", keyword_text) if part]
            if not parts:
                return False
            pattern = r"(?<![a-z0-9])" + r"[\s\-_]*".join(parts) + r"(?![a-z0-9])"
            return bool(re.search(pattern, clue_text))
        normalized_keyword = self._blr_normalize_clue_text(keyword_text)
        return bool(normalized_keyword) and normalized_keyword in self._blr_normalize_clue_text(clue_text)

    def _blr_any_clue_matches(self, clues, keywords):
        return any(self._blr_clue_contains_keyword(clue, keyword) for clue in clues for keyword in keywords)

    def _blr_has_fixed_asset_lines(self, strict=False):
        self.ensure_one()
        invoice_lines = self._blr_relevant_invoice_lines()
        if not invoice_lines:
            return None if strict else False
        asset_flags = {getattr(line.account_id, "account_type", "") == "asset_fixed" for line in invoice_lines}
        if strict and len(asset_flags) != 1:
            return None
        return any(asset_flags)

    def _blr_guess_tax_type(self):
        self.ensure_one()
        fmt = (self.tw_blr_format_code or "").strip()
        clues = self._blr_tax_clues()

        if fmt == "29":
            return "1", False
        if fmt == "28":
            return ("3" if self._blr_tax_amount_is_zero() else "1"), False
        if not self._blr_tax_amount_is_zero():
            return "1", False

        has_exempt = self._blr_any_clue_matches(clues, EXEMPT_KEYWORDS)
        has_zero_rate = self._blr_any_clue_matches(clues, ZERO_RATE_KEYWORDS)
        if has_exempt and not has_zero_rate:
            return "3", False
        if has_zero_rate and not has_exempt:
            return "2", False
        if not clues:
            return False, _("找不到 Odoo 稅別或稅務標籤資料，無法區分零稅率或免稅，請手動確認課稅別。")
        return False, _("Odoo 現有稅別或稅務標籤名稱不足以明確區分零稅率或免稅，請手動確認課稅別。")

    def _blr_guess_deduct_code(self, tax_type):
        self.ensure_one()
        if self.move_type not in ("in_invoice", "in_refund"):
            return False, False
        if tax_type not in ("1", "2", "3"):
            return False, _("請先確認課稅別後再自動帶入抵扣代號。")

        is_asset = self._blr_has_fixed_asset_lines(strict=True)
        if is_asset is None:
            return False, _("同一張進項憑證同時包含固定資產與進貨費用，無法自動判定抵扣代號，請手動確認。")

        deductible_code = "2" if is_asset else "1"
        nondeductible_code = "4" if is_asset else "3"
        if tax_type in ("2", "3"):
            return nondeductible_code, False

        clues = self._blr_tax_clues()
        has_nondeductible = self._blr_any_clue_matches(clues, NON_DEDUCTIBLE_KEYWORDS)
        has_deductible = self._blr_any_clue_matches(clues, DEDUCTIBLE_KEYWORDS)
        if has_nondeductible and not has_deductible:
            return nondeductible_code, False
        if has_deductible and not has_nondeductible:
            return deductible_code, False
        if not clues:
            return False, _("找不到 Odoo 稅別或稅務標籤資料，無法自動判定抵扣代號，請手動確認。")
        return False, _("Odoo 現有稅別或稅務標籤資料不足以判定抵扣代號是否可扣抵，請手動確認。")

    def _blr_get_export_deduct_code(self):
        self.ensure_one()
        if self.move_type not in ("in_invoice", "in_refund"):
            return " "
        return (self.tw_deduct_code or "")[:1]

    @api.onchange("tw_blr_invoice_autofill_trigger")
    def _onchange_tw_blr_invoice_autofill_trigger(self):
        warnings = []
        for move in self:
            if move.tw_blr_invoice_autofill_trigger:
                message = move._blr_apply_invoice_autofill()
                if message:
                    warnings.append(message)
            move.tw_blr_invoice_autofill_trigger = False
        return self._blr_build_onchange_warning(warnings)

    @api.onchange("tw_blr_tax_autofill_trigger")
    def _onchange_tw_blr_tax_autofill_trigger(self):
        warnings = []
        for move in self:
            if move.tw_blr_tax_autofill_trigger:
                tax_type, tax_warning = move._blr_guess_tax_type()
                move.tw_tax_type = tax_type or False
                if tax_warning:
                    warnings.append(tax_warning)

                if move.move_type in ("out_invoice", "out_refund"):
                    move.tw_deduct_code = False
                elif tax_type:
                    deduct_code, deduct_warning = move._blr_guess_deduct_code(tax_type)
                    move.tw_deduct_code = deduct_code or False
                    if deduct_warning:
                        warnings.append(deduct_warning)
                else:
                    move.tw_deduct_code = False
            move.tw_blr_tax_autofill_trigger = False
        return self._blr_build_onchange_warning(warnings)

    @api.onchange("move_type")
    def _onchange_move_type_blr_scope(self):
        for move in self:
            fmt = move.tw_blr_format_code or ""
            if move.move_type in ("in_invoice", "in_refund") and fmt not in IN_FORMAT_CODES:
                move.tw_blr_format_code = False
            if move.move_type in ("out_invoice", "out_refund") and fmt not in OUT_FORMAT_CODES:
                move.tw_blr_format_code = False
                move.tw_deduct_code = False

    @api.onchange("tw_blr_format_code")
    def _onchange_tw_blr_format_code(self):
        for move in self:
            fmt = move.tw_blr_format_code or ""
            if fmt.startswith("3"):
                move.tw_deduct_code = False
            if fmt == "29":
                move.tw_tax_type = "1"
            elif fmt == "28" and move.tw_tax_type not in ("1", "3"):
                move.tw_tax_type = "1"
            if fmt in CUSTOMS_FORMATS:
                move.tw_invoice_track = False
                move.tw_invoice_number = False
                move.tw_other_voucher_no = False
                move.tw_utility_carrier_no = False
            else:
                move.tw_customs_pay_no = False
            if fmt in INV_OR_UTILITY_FORMATS:
                move.tw_other_voucher_no = False
            if fmt in INV_OR_OTHER_FORMATS:
                move.tw_utility_carrier_no = False
            if fmt in INV_ONLY_FORMATS:
                move.tw_other_voucher_no = False
                move.tw_utility_carrier_no = False
                move.tw_customs_pay_no = False

    @api.onchange("tw_utility_carrier_no")
    def _onchange_tw_utility_carrier_no(self):
        for move in self:
            if move.tw_utility_carrier_no:
                move.tw_invoice_track = False
                move.tw_invoice_number = False

    @api.onchange("tw_other_voucher_no")
    def _onchange_tw_other_voucher_no(self):
        for move in self:
            if move.tw_other_voucher_no:
                move.tw_invoice_track = False
                move.tw_invoice_number = False

    @api.onchange("tw_invoice_track", "tw_invoice_number")
    def _onchange_tw_invoice_identifier(self):
        for move in self:
            if move.tw_invoice_track or move.tw_invoice_number:
                move.tw_other_voucher_no = False
                move.tw_utility_carrier_no = False

    @api.onchange("tw_customs_pay_no")
    def _onchange_tw_customs_pay_no(self):
        for move in self:
            if move.tw_customs_pay_no:
                move.tw_invoice_track = False
                move.tw_invoice_number = False
                move.tw_other_voucher_no = False
                move.tw_utility_carrier_no = False

    @api.onchange("tw_tax_type", "tw_deduct_code")
    def _onchange_tw_tax_deduct_cross(self):
        for move in self:
            fmt = move.tw_blr_format_code or ""
            if not fmt or fmt.startswith("3"):
                continue
            if move.tw_tax_type in ("2", "3") and move.tw_deduct_code in ("1", "2"):
                move.tw_deduct_code = "3"
            if move.tw_deduct_code in ("1", "2") and move.tw_tax_type != "1":
                move.tw_tax_type = "1"

    def action_post(self):
        self._blr_pre_post_validate()
        return super().action_post()

    def _blr_pre_post_validate(self):
        for move in self:
            if move.tw_blr_skip_export:
                continue
            if move.move_type not in ("out_invoice", "out_refund", "in_invoice", "in_refund"):
                continue
            if not move.tw_blr_format_code:
                continue
            move._blr_validate_fields_raise()

    def _blr_validate_fields_raise(self):
        self.ensure_one()
        fmt = (self.tw_blr_format_code or "").strip()

        def _err(msg):
            raise ValidationError(_("%s：%s") % (self.display_name, msg))

        def _is_track(value):
            return bool(re.fullmatch(r"[A-Z]{2}", (value or "").strip().upper()))

        def _is_invno(value):
            return bool(re.fullmatch(r"\d{8}", (value or "").strip()))

        def _is_other(value):
            return bool(re.fullmatch(r"[0-9A-Z]{10}", (value or "").strip().upper()))

        def _is_utility(value):
            return bool(re.fullmatch(r"BB[0-9A-Z]{8}", (value or "").strip().upper()))

        def _is_customs(value):
            return bool(re.fullmatch(r"[0-9A-Z]{14}", (value or "").strip().upper()))

        if len(fmt) != 2:
            _err(_("申報格式代號必須為2碼"))
        if fmt in UNSUPPORTED_FORMATS:
            _err(_("目前版本未開放特種稅額格式代號%s" % fmt))
        if self.move_type in ("in_invoice", "in_refund") and fmt not in IN_FORMAT_CODES:
            _err(_("採購/進項單據只允許使用21～29的申報格式代號"))
        if self.move_type in ("out_invoice", "out_refund") and fmt not in OUT_FORMAT_CODES:
            _err(_("銷售/銷項單據只允許使用31～36的申報格式代號"))
        if self.tw_tax_type not in ("1", "2", "3"):
            _err(_("課稅別必須填寫1/2/3"))

        track_ok = _is_track(self.tw_invoice_track)
        inv_ok = _is_invno(self.tw_invoice_number)
        other_ok = _is_other(self.tw_other_voucher_no)
        utility_ok = _is_utility(self.tw_utility_carrier_no)
        customs_ok = _is_customs(self.tw_customs_pay_no)
        has_inv = bool((self.tw_invoice_track or "").strip() or (self.tw_invoice_number or "").strip())
        has_other = bool((self.tw_other_voucher_no or "").strip())
        has_utility = bool((self.tw_utility_carrier_no or "").strip())
        has_customs = bool((self.tw_customs_pay_no or "").strip())

        if has_inv and not track_ok:
            _err(_("發票字軌需為2碼大寫英文字母"))
        if has_inv and not inv_ok:
            _err(_("發票號碼需為8碼數字"))
        if has_other and not other_ok:
            _err(_("其他憑證號碼需為10碼英數"))
        if has_utility and not utility_ok:
            _err(_("公用事業載具流水號需為BB開頭加8碼英數"))
        if has_customs and not customs_ok:
            _err(_("海關代徵營業稅繳納證號碼需為14碼英數"))

        if fmt.startswith("3"):
            if self.tw_deduct_code:
                _err(_("銷項單據不可填抵扣代號"))
        else:
            if self.tw_deduct_code not in ("1", "2", "3", "4"):
                _err(_("進項憑證必須填寫抵扣代號1/2/3/4"))
            asset_flag = self._blr_has_fixed_asset_lines(strict=True)
            if asset_flag is True and self.tw_deduct_code not in ("2", "4"):
                _err(_("固定資產進項憑證的抵扣代號必須為2或4"))
            elif asset_flag is False and self.tw_deduct_code not in ("1", "3"):
                _err(_("進貨及費用的抵扣代號必須為1或3"))
            if self.tw_tax_type in ("2", "3") and self.tw_deduct_code in ("1", "2"):
                _err(_("課稅別為零稅率或免稅時，抵扣代號不可為1或2"))
            if self.tw_deduct_code in ("1", "2") and self.tw_tax_type != "1":
                _err(_("抵扣代號為1或2時，課稅別必須為1"))

        if fmt == "29" and self.tw_tax_type != "1":
            _err(_("格式29課稅別只允許1 應稅"))
        if fmt == "28" and self.tw_tax_type not in ("1", "3"):
            _err(_("格式28課稅別只允許1 應稅或3 免稅"))

        if fmt in CUSTOMS_FORMATS:
            if not customs_ok:
                _err(_("格式%s必須填寫海關代徵營業稅繳納證號碼" % fmt))
            if has_inv or has_other or has_utility:
                _err(_("格式%s不可同時填寫發票或其他憑證欄位" % fmt))
            return

        if fmt in INV_OR_UTILITY_FORMATS:
            if utility_ok and (has_inv or has_other or has_customs):
                _err(_("格式%s不可同時填寫公用事業載具流水號與其他識別欄位" % fmt))
            if utility_ok:
                return
            if track_ok and inv_ok and not has_other and not has_customs:
                return
            _err(_("格式%s需擇一填寫發票字軌+發票號碼，或公用事業載具流水號" % fmt))

        if fmt in INV_OR_OTHER_FORMATS:
            if other_ok and (has_inv or has_utility or has_customs):
                _err(_("格式%s不可同時填寫其他憑證號碼與其他識別欄位" % fmt))
            if other_ok:
                return
            if track_ok and inv_ok and not has_utility and not has_customs:
                return
            _err(_("格式%s需擇一填寫發票字軌+發票號碼，或其他憑證號碼" % fmt))

        if fmt in INV_ONLY_FORMATS:
            if not (track_ok and inv_ok):
                _err(_("格式%s需填寫發票字軌(2碼)+發票號碼(8碼)" % fmt))
            if has_other or has_utility or has_customs:
                _err(_("格式%s不可填寫其他憑證號碼、公用事業載具流水號或海關號碼" % fmt))