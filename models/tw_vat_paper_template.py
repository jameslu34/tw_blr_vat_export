import json

from odoo import api, fields, models


BUILTIN_PAPER_PAGE_MAP = {
    "401": 0,
    "403": 2,
    "404": 4,
}


COMMON_HEADER_FIELDS = {
    "vat_no": {"x": 186, "y": 523, "size": 10, "align": "right"},
    "tax_id_9": {"x": 213, "y": 491, "size": 10, "align": "right"},
    "roc_year": {"x": 186, "y": 523, "size": 10, "align": "right"},
    "period_month": {"x": 213, "y": 491, "size": 10, "align": "right"},
}


def _legacy_vat_form_layout_dict():
    return {
        "page": 0,
        "grid_step": 20,
        "fields": {
            "roc_year": {"x": 145, "y": 781, "size": 11},
            "period_month": {"x": 214, "y": 781, "size": 11},
            "filing_code": {"x": 283, "y": 781, "size": 11},
            "vat_no": {"x": 520, "y": 781, "size": 11, "align": "right"},
            "tax_id_9": {"x": 520, "y": 760, "size": 11, "align": "right"},
            "taxable_sales": {"x": 520, "y": 525, "size": 11, "align": "right"},
            "output_tax": {"x": 520, "y": 509, "size": 11, "align": "right"},
            "input_tax_deductible": {"x": 520, "y": 290, "size": 11, "align": "right"},
            "net_tax_payable": {"x": 520, "y": 165, "size": 11, "align": "right"},
            "net_tax_credit": {"x": 520, "y": 149, "size": 11, "align": "right"},
        },
    }


def _vat_form_layout_dict(form_type=None):
    if form_type == "401":
        return {
            "page": BUILTIN_PAPER_PAGE_MAP["401"],
            "grid_step": 20,
            "fields": {
                **COMMON_HEADER_FIELDS,
                "taxable_sales": {"x": 492, "y": 348, "size": 10, "align": "right"},
                "output_tax": {"x": 733, "y": 445, "size": 10, "align": "right"},
                "input_tax_deductible": {"x": 733, "y": 429, "size": 10, "align": "right"},
                "net_tax_payable": {"x": 733, "y": 382, "size": 10, "align": "right"},
                "net_tax_credit": {"x": 733, "y": 366, "size": 10, "align": "right"},
            },
        }
    if form_type == "403":
        return {
            "page": BUILTIN_PAPER_PAGE_MAP["403"],
            "grid_step": 20,
            "fields": dict(COMMON_HEADER_FIELDS),
        }
    if form_type == "404":
        return {
            "page": BUILTIN_PAPER_PAGE_MAP["404"],
            "grid_step": 20,
            "fields": dict(COMMON_HEADER_FIELDS),
        }
    return _legacy_vat_form_layout_dict()


def _default_vat_form_layout(form_type=None):
    return json.dumps(
        _vat_form_layout_dict(form_type),
        ensure_ascii=False,
        indent=2,
    )


class TwVatPaperTemplate(models.Model):
    _name = "tw.vat.paper.template"
    _description = "臺灣營業稅紙本申報書相容模板"
    _order = "company_id, form_type, id"

    company_id = fields.Many2one("res.company", required=True, ondelete="cascade", string="公司")
    form_type = fields.Selection(
        [("401", "401"), ("403", "403"), ("404", "404")],
        required=True,
        default="401",
        string="申報書類型",
    )
    template = fields.Binary(string="相容模板")
    template_name = fields.Char(string="模板檔名")
    layout = fields.Text(string="相容座標(JSON)", default=lambda self: self._default_layout_for_context())

    _sql_constraints = [
        (
            "tw_vat_paper_template_company_form_type_uniq",
            "unique(company_id, form_type)",
            "同一公司每種申報書類型只能設定一份模板。",
        )
    ]

    @api.model
    def _default_layout_for_context(self):
        return _default_vat_form_layout(self.env.context.get("default_form_type"))

    @api.model
    def _get_default_layout_data(self, form_type=None):
        return _vat_form_layout_dict(form_type)

    @api.model
    def _get_legacy_default_layout_data(self):
        return _legacy_vat_form_layout_dict()

    @api.model
    def _get_builtin_template_page(self, form_type):
        return BUILTIN_PAPER_PAGE_MAP.get(form_type, 0)

    @api.onchange("form_type")
    def _onchange_form_type(self):
        for record in self:
            if not record.form_type:
                continue
            if not record.layout:
                record.layout = _default_vat_form_layout(record.form_type)
                continue
            try:
                current_layout = json.loads(record.layout)
            except Exception:
                continue
            if current_layout in (_legacy_vat_form_layout_dict(), _vat_form_layout_dict(record.form_type)):
                record.layout = _default_vat_form_layout(record.form_type)

    def name_get(self):
        return [(record.id, f"{record.form_type}申報書相容模板") for record in self]