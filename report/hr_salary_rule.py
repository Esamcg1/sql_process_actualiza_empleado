# -*- coding: utf-8 -*-

from odoo import api, models, fields

class HrSalary_rule(models.Model):
    _inherit = 'hr.salary.rule'
    _description = 'hr.salary.rule'

    x_code = fields.Char(string="Codigo personalizado", help="Codigo diferente a (code) para modulo sql_process")

