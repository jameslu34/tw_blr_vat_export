from odoo import api, fields, models


_PARAM_FIELDS = (
    "tw_tax_id_9",
    "tw_responsible_name",
    "tw_filer_idno",
    "tw_filer_name",
    "tw_filer_tel_area",
    "tw_filer_tel",
    "tw_filer_tel_ext",
    "tw_agent_reg_no",
)


class ResCompany(models.Model):
    _inherit = "res.company"

    tw_tax_id_9 = fields.Char(
        string="稅籍編號(9碼)",
        compute="_compute_tw_tax_id_9",
        inverse="_inverse_tw_tax_id_9",
        store=False,
    )
    tw_responsible_name = fields.Char(
        string="負責人姓名",
        compute="_compute_tw_responsible_name",
        inverse="_inverse_tw_responsible_name",
        store=False,
    )
    tw_filer_idno = fields.Char(
        string="申報人身分證統一編號(10碼，可留白)",
        compute="_compute_tw_filer_idno",
        inverse="_inverse_tw_filer_idno",
        store=False,
    )
    tw_filer_name = fields.Char(
        string="申報人姓名",
        compute="_compute_tw_filer_name",
        inverse="_inverse_tw_filer_name",
        store=False,
    )
    tw_filer_tel_area = fields.Char(
        string="申報人電話區碼",
        compute="_compute_tw_filer_tel_area",
        inverse="_inverse_tw_filer_tel_area",
        store=False,
    )
    tw_filer_tel = fields.Char(
        string="申報人電話",
        compute="_compute_tw_filer_tel",
        inverse="_inverse_tw_filer_tel",
        store=False,
    )
    tw_filer_tel_ext = fields.Char(
        string="申報人電話分機",
        compute="_compute_tw_filer_tel_ext",
        inverse="_inverse_tw_filer_tel_ext",
        store=False,
    )
    tw_agent_reg_no = fields.Char(
        string="代理申報人登錄(文)字號(委任申報時填寫)",
        compute="_compute_tw_agent_reg_no",
        inverse="_inverse_tw_agent_reg_no",
        store=False,
    )
    tw_vat_paper_template_ids = fields.One2many(
        "tw.vat.paper.template",
        "company_id",
        string="紙本申報書模板(相容欄位)",
    )

    def _tw_param_key(self, field_name):
        self.ensure_one()
        return f"tw_blr_vat_export.{field_name}.{self.id}"

    @api.model
    def _tw_legacy_company_columns(self):
        self.env.cr.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = current_schema()
              AND table_name = %s
              AND column_name = ANY(%s)
            """,
            [self._table, list(_PARAM_FIELDS)],
        )
        return {row[0] for row in self.env.cr.fetchall()}

    def _tw_get_legacy_company_values(self, field_name):
        if not self.ids or field_name not in self._tw_legacy_company_columns():
            return {}
        self.env.cr.execute(
            f'SELECT id, "{field_name}" FROM {self._table} WHERE id = ANY(%s)',
            [list(self.ids)],
        )
        return {row[0]: row[1] for row in self.env.cr.fetchall()}

    def _compute_tw_param_field(self, field_name):
        icp = self.env["ir.config_parameter"].sudo()
        legacy_values = self._tw_get_legacy_company_values(field_name)
        for company in self:
            value = icp.get_param(company._tw_param_key(field_name), default=None)
            if value in (None, ""):
                value = legacy_values.get(company.id) or ""
            company[field_name] = value

    def _inverse_tw_param_field(self, field_name):
        icp = self.env["ir.config_parameter"].sudo()
        for company in self:
            icp.set_param(company._tw_param_key(field_name), company[field_name] or "")

    @api.depends_context("uid")
    def _compute_tw_tax_id_9(self):
        self._compute_tw_param_field("tw_tax_id_9")

    def _inverse_tw_tax_id_9(self):
        self._inverse_tw_param_field("tw_tax_id_9")

    @api.depends_context("uid")
    def _compute_tw_responsible_name(self):
        self._compute_tw_param_field("tw_responsible_name")

    def _inverse_tw_responsible_name(self):
        self._inverse_tw_param_field("tw_responsible_name")

    @api.depends_context("uid")
    def _compute_tw_filer_idno(self):
        self._compute_tw_param_field("tw_filer_idno")

    def _inverse_tw_filer_idno(self):
        self._inverse_tw_param_field("tw_filer_idno")

    @api.depends_context("uid")
    def _compute_tw_filer_name(self):
        self._compute_tw_param_field("tw_filer_name")

    def _inverse_tw_filer_name(self):
        self._inverse_tw_param_field("tw_filer_name")

    @api.depends_context("uid")
    def _compute_tw_filer_tel_area(self):
        self._compute_tw_param_field("tw_filer_tel_area")

    def _inverse_tw_filer_tel_area(self):
        self._inverse_tw_param_field("tw_filer_tel_area")

    @api.depends_context("uid")
    def _compute_tw_filer_tel(self):
        self._compute_tw_param_field("tw_filer_tel")

    def _inverse_tw_filer_tel(self):
        self._inverse_tw_param_field("tw_filer_tel")

    @api.depends_context("uid")
    def _compute_tw_filer_tel_ext(self):
        self._compute_tw_param_field("tw_filer_tel_ext")

    def _inverse_tw_filer_tel_ext(self):
        self._inverse_tw_param_field("tw_filer_tel_ext")

    @api.depends_context("uid")
    def _compute_tw_agent_reg_no(self):
        self._compute_tw_param_field("tw_agent_reg_no")

    def _inverse_tw_agent_reg_no(self):
        self._inverse_tw_param_field("tw_agent_reg_no")
