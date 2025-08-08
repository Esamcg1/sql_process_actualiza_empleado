# -*- coding: utf-8 -*-
from tempfile import NamedTemporaryFile
from openpyxl import Workbook
from openpyxl.styles import Color, Fill, Font, Alignment
from openpyxl.cell import Cell
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from datetime import datetime
import base64
from odoo import api, models, fields

class GenerateStockWizard(models.Model):
    _name = "sql_process_actualiza.sql_actualiza_execute"

    data = fields.Binary('File', readonly=True)
    name = fields.Char('File Name', readonly=True)
    state = fields.Selection([('choose', 'choose'), ('get', 'get')], default='choose')

    categoria_employee_id = fields.Many2one('hr.employee.category', string='Tipo de Planilla')
    imprime = fields.Boolean(string='¿Imprime?', default=True, required=True)
    actualiza = fields.Boolean(string='¿Actualiza?', default=True, required=True)
    employee_id = fields.Many2one('hr.employee', string='Empleado')
    fecha_desde = fields.Date(string="Fecha Desde", default=fields.Date.today, required=True)
    fecha_hasta = fields.Date(string="Fecha Hasta", default=fields.Date.today, required=True)

    def generate_file(self):
        # Si debe actualizar datos
        if self.actualiza:
            self._actualizar_datos_empleado()

        # Crear libro de Excel
        wb = Workbook()

        # Configurar estilos base
        al1 = Alignment(horizontal="center", vertical="center")
        font = Font(size=12, bold=True, color="FFFFFF")
        fill = PatternFill("solid", fgColor="4472C4")
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)

        # Generar hojas
        self._generar_resumen_empleado(wb)
        self._generar_detalle_empleado(wb, border, al1, font, fill)

        # Guardar archivo temporal
        tmp = NamedTemporaryFile(suffix=".xlsx", delete=False)
        wb.save(tmp.name)
        tmp.seek(0)
        excel_data = tmp.read()
        tmp.close()

        # Guardar en campos del wizard
        self.write({
            'data': base64.b64encode(excel_data),
            'name': 'actualiza_empleado.xlsx',
            'state': 'get'
        })

        # Volver a mostrar wizard
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'new',
        }


    # --------------------------------------------------------
    # -------------- Actualiza Empleado ---------------------
    # --------------------------------------------------------
    def _actualizar_datos_empleado(self):
        """
        Actualiza los campos relacionados con días laborados y salario anual en los empleados,
        según el rango de fechas y si es un empleado específico o todos.
        """
        if self.employee_id:
            # Consulta para empleado específico
            sql = self._get_sql_update_empleado(especifico=True)
            params = (
                self.employee_id.id, self.fecha_desde, self.fecha_hasta, self.fecha_desde,
                self.employee_id.id, self.fecha_desde, self.fecha_hasta, self.fecha_desde
            )
        else:
            # Consulta para todos los empleados
            sql = self._get_sql_update_empleado(especifico=False)
            params = (
                self.fecha_desde, self.fecha_hasta, self.fecha_desde,
                self.fecha_desde, self.fecha_hasta, self.fecha_desde
            )

        self.env.cr.execute(sql, params)

    def _get_sql_update_empleado(self, especifico=False):
        """
        Devuelve el SQL para actualizar días laborados, salario anual y bonificación.
        :param especifico: True si es para un empleado específico, False para todos.
        """
        where_empleado = "he.id = %s" if especifico else "he.id >= 1"

        return f"""
            WITH UltimoSalario AS (
                SELECT 
                    he.id AS employee_id,
                    COALESCE(hpl.total, 0) AS ultimo_salario,
                    ROW_NUMBER() OVER (
                        PARTITION BY he.id 
                        ORDER BY hpr.date_start DESC, hpr.id DESC
                    ) AS rn
                FROM hr_payslip_run hpr
                JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                JOIN hr_contract hc ON hc.id = hp.contract_id
                JOIN res_company rc ON rc.id = hp.company_id
                JOIN hr_employee he ON hp.employee_id = he.id
                LEFT JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                LEFT JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                WHERE {where_empleado}
                AND hsr.x_code IN ('SO', 'SO2')
                AND hpr.date_start >= %s 
                AND hpr.date_start <= %s
                AND hpr.date_end >= hc.date_start 
                AND (hc.date_end >= %s OR hc.date_end IS NULL)
            ),
            CalculatedData AS (
                SELECT 
                    he.id AS employee_id,
                    COALESCE(SUM(CASE WHEN hsr.x_code = 'raDLA' THEN hpl.total ELSE 0 END), 0) AS dias,
                    COALESCE(SUM(CASE WHEN hsr.x_code IN ('ORDQ1', 'ORDQ2') THEN hpl.total ELSE 0 END), 0) AS salario_ordinario_anual,
                    COALESCE(MAX(hc.otra_bonificacion), 0) AS bonificacion_decreto,
                    COALESCE(us.ultimo_salario, 0) AS ultimo_salario
                FROM hr_payslip_run hpr
                JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                JOIN hr_contract hc ON hc.id = hp.contract_id
                JOIN res_company rc ON rc.id = hp.company_id
                JOIN hr_employee he ON hp.employee_id = he.id
                LEFT JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                LEFT JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                LEFT JOIN UltimoSalario us ON us.employee_id = he.id AND us.rn = 1
                WHERE {where_empleado}
                AND hsr.x_code IN ('raDLA', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14', 'SO', 'SO2')
                AND hpr.date_start >= %s 
                AND hpr.date_start <= %s
                AND hpr.date_end >= hc.date_start 
                AND (hc.date_end >= %s OR hc.date_end IS NULL)
                GROUP BY he.id, us.ultimo_salario
            )
            UPDATE hr_employee em
            SET 
                dias_laborados_ano = cd.dias,
                salario_anual_nominal = cd.salario_ordinario_anual,
                bonificacion_decreto = cd.bonificacion_decreto,
                x_ultimo_salario_ordinario = cd.ultimo_salario
            FROM CalculatedData cd
            WHERE em.id = cd.employee_id;
        """

#--------------------------------------------------------#
# ---------- Resumen por Empleado -----------------------#
#--------------------------------------------------------#

    def _generar_resumen_empleado(self, wb):
        """
        Genera la hoja 'Resumen Empleado' en el Excel con la información resumida por empleado.
        """
        ws = wb.active
        ws.title = "Resumen Empleado"

        # Configurar hoja y estilos
        self._configurar_hoja_resumen(ws)

        # Ejecutar SQL y obtener datos
        datos = self._obtener_datos_resumen_empleado()

        # Llenar datos en la hoja
        self._llenar_datos_resumen(ws, datos)
        

    def _configurar_hoja_resumen(self, ws):
        """
        Configura la hoja de resumen: orientación, tamaño, columnas y cabeceras.
        """
        from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToHeight = 0
        ws.page_setup.fitToWidth = 1

        al1 = Alignment(horizontal="left", vertical="center")
        font = Font(size=12, bold=True, italic=True, color="E0D7D7")
        fill = PatternFill("solid", fgColor="4777AD")
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)

        self.style_range(ws, 'A1:J1', border=border, alignment=al1, font=font, fill=fill)

        # Anchos de columna
        widths = [40, 40, 13, 14, 8, 16, 16, 16, 16, 16]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = width

        # Cabeceras
        headers = [
            "Empleado", "Departamento", "Fecha Contrato", "Salario Actual",
            "Días", "Salario", "Bonificación", "Ordinario", "Otros Ingresos", "Sueldo"
        ]
        for col, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col).value = header


    def _obtener_datos_resumen_empleado(self):
        """
        Ejecuta el SQL que genera el resumen por empleado.
        """

        company_id = self.env.user.company_id.id
        sql_base = """
            SELECT r.contract_id, he.id AS employee_id, he.name AS nombre_empleado, he.department_id,
                CONCAT(
                    COALESCE(hda.name->>'es',''), 
                    (CASE WHEN hda.name IS NULL THEN '' ELSE '/' END),
                    COALESCE(hdp.name->>'es',''), 
                    (CASE WHEN hdp.name IS NULL THEN '' ELSE '/' END),
                    COALESCE(hd.name->>'es','')
                ) AS departamento,
                COALESCE(he.bonificacion_decreto, 0) AS bonificacion_decreto, 
                COALESCE(he.bonificaciones, 0) AS bonificaciones, 
                hc.date_start AS fecha_inicio_contrato,
                COALESCE(hc.ingresos1, 0) AS ingresos1, 
                COALESCE(hc.ingresos4, 0) AS ingresos4, 
                COALESCE(r.dias, 0) AS dias, 
                COALESCE(hc.wage, 0) AS salario_actual,
                COALESCE(r.bonificacion_acumulada, 0) AS bonificacion_acumulada, 
                COALESCE(r.sueldo_acumulado, 0) AS sueldo_acumulado, 
                COALESCE(r.salario_ordinario_anual, 0) AS salario_ordinario_anual,
                COALESCE(hc.ingresos1, 0) + COALESCE(hc.ingresos4, 0) AS bonificacion_adicional
            FROM (
                SELECT d.contract_id,
                    COALESCE(SUM(d.dias), 0) AS dias,
                    COALESCE(SUM(d.ordinario_quincena1), 0) + COALESCE(SUM(d.ordinario_quincena2), 0) AS salario_ordinario_anual,
                    COALESCE(SUM(d.bonificacion), 0) AS bonificacion_acumulada,
                    COALESCE(SUM(d.bonificacion), 0) AS sueldo_acumulado
                FROM (
                    SELECT hp.contract_id,
                        SUM(CASE WHEN hsr.x_code = 'raDLA' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS dias,
                        SUM(CASE WHEN hsr.x_code = 'SO' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS salario_mensual,
                        SUM(CASE WHEN hsr.x_code = 'ORDQ1' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS ordinario_quincena1,
                        SUM(CASE WHEN hsr.x_code = 'ORDQ2' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS ordinario_quincena2,
                        SUM(CASE WHEN hsr.x_code = 'LiqAG' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS aguinaldo,
                        SUM(CASE WHEN hsr.x_code = 'SO' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS bonificacion, 
                        SUM(CASE WHEN hsr.x_code = 'LiqB14' THEN COALESCE(hpl.total, 0) ELSE 0 END) AS bono_14
                    FROM hr_payslip_run hpr
                    JOIN hr_payslip hp ON hp.payslip_run_id = hpr.id
                    JOIN hr_contract hc ON hp.contract_id = hc.id
                    JOIN res_company rc ON hp.company_id = rc.id                  
                    LEFT JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                    LEFT JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                    WHERE rc.id = %s
        """
        params = []
        if self.employee_id:
            sql_base += """
                    AND hp.employee_id = %s
                    AND hsr.x_code IN ('raDLA', 'SO', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14')
                    AND hpr.date_start >= %s AND hpr.date_start <= %s
                    AND hpr.date_end >= hc.date_start AND hc.date_end IS NULL
                    GROUP BY hp.contract_id, hsr.x_code
                ) d 
                GROUP BY d.contract_id
            ) r
            JOIN hr_contract hc ON r.contract_id = hc.id
            LEFT JOIN hr_employee he ON he.id = hc.employee_id
            LEFT JOIN hr_department hd ON he.department_id = hd.id
            LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
            LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
            ORDER BY he.name
            """
            params = [company_id, self.employee_id.id, self.fecha_desde, self.fecha_hasta]
        else:
            sql_base += """
                    AND hp.employee_id >= 1
                    AND hsr.x_code IN ('raDLA', 'SO', 'ORDQ1', 'ORDQ2', 'LiqAG', 'LiqB14')
                    AND hpr.date_start >= %s AND hpr.date_start <= %s
                    AND hpr.date_end >= hc.date_start AND hc.date_end IS NULL
                    GROUP BY hp.contract_id, hsr.x_code
                ) d 
                GROUP BY d.contract_id
            ) r
            JOIN hr_contract hc ON r.contract_id = hc.id
            LEFT JOIN hr_employee he ON he.id = hc.employee_id
            LEFT JOIN hr_department hd ON he.department_id = hd.id
            LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
            LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
            ORDER BY he.name
            """
            params = [company_id, self.fecha_desde, self.fecha_hasta]

        self.env.cr.execute(sql_base, params)
        return self.env.cr.dictfetchall()


    def _llenar_datos_resumen(self, ws, datos):
        """
        Llena la hoja de resumen con los datos y calcula totales.
        """
        from openpyxl.styles import Font

        row = 2
        wtotal_salario = 0
        wtotal_bonificacion = 0
        wtotal_ordinario = 0
        wtotal_otros_ingresos = 0
        wtotal_sueldo = 0

        for line in datos:
            ws.cell(row=row, column=1).value = line['nombre_empleado']
            ws.cell(row=row, column=2).value = line['departamento']

            if line['fecha_inicio_contrato']:
                fs = line['fecha_inicio_contrato'].strftime('%Y-%m-%d').split('-')
                ws.cell(row=row, column=3).value = f"{fs[2]}/{fs[1]}/{fs[0]}"

            ws.cell(row=row, column=4).value = round(line['salario_actual'] or 0, 0)
            ws.cell(row=row, column=4).number_format = '#,##0.00'

            ws.cell(row=row, column=5).value = round(line['dias'] or 0, 0)
            ws.cell(row=row, column=5).number_format = '#,##0'

            ws.cell(row=row, column=6).value = round(line['salario_ordinario_anual'] or 0, 2)
            ws.cell(row=row, column=6).number_format = '#,##0.00'

            ws.cell(row=row, column=7).value = round(line['bonificacion_adicional'] or 0, 2)
            ws.cell(row=row, column=7).number_format = '#,##0.00'

            ws.cell(row=row, column=8).value = round(line['bonificacion_decreto'] or 0, 2)
            ws.cell(row=row, column=8).number_format = '#,##0.00'

            ws.cell(row=row, column=9).value = round(line['bonificacion_acumulada'] or 0, 2)
            ws.cell(row=row, column=9).number_format = '#,##0.00'

            ws.cell(row=row, column=10).value = round(line['sueldo_acumulado'] or 0, 2)
            ws.cell(row=row, column=10).number_format = '#,##0.00'

            wtotal_salario += line['salario_ordinario_anual'] or 0
            wtotal_bonificacion += line['bonificacion_adicional'] or 0
            wtotal_ordinario += line['bonificacion_decreto'] or 0
            wtotal_otros_ingresos += line['bonificacion_acumulada'] or 0
            wtotal_sueldo += line['sueldo_acumulado'] or 0

            row += 1

        # Totales
        ws.cell(row=row, column=1).value = 'Total: '
        ws.cell(row=row, column=1).font = Font(size=13, bold=True)

        totales = [wtotal_salario, wtotal_bonificacion, wtotal_ordinario, wtotal_otros_ingresos, wtotal_sueldo]
        for i, total in enumerate(totales, start=6):
            ws.cell(row=row, column=i).value = total
            ws.cell(row=row, column=i).number_format = '#,##0.00'
            ws.cell(row=row, column=i).font = Font(size=13, bold=True)

                
#--------------------------------------------------------#
# ---------- Detalle por Empleado -----------------------#
#--------------------------------------------------------#
    def _generar_detalle_empleado(self, wb, border, al1, font, fill):
        """
        Genera la hoja 'Detalle Empleado' con el detalle de cada nómina por empleado.
        """
        ws = wb.create_sheet(title="Detalle Empleado")

        # Configurar hoja y estilos
        self._configurar_hoja_detalle(ws, border, al1, font, fill)

        # Ejecutar SQL y obtener datos
        datos = self._obtener_datos_detalle_empleado()

        # Llenar datos en la hoja
        self._llenar_datos_detalle(ws, datos)


    def _configurar_hoja_detalle(self, ws, border, al1, font, fill):
        """
        Configura la hoja del detalle: ancho columnas y cabeceras.
        """
        self.style_range(ws, 'A1:L1', border=border, alignment=al1, font=font, fill=fill)

        widths = [40, 44, 40, 13, 13, 13, 8, 16, 16, 16, 16, 16]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = width

        headers = [
            "Empleado", "Nómina", "Departamento", "Inicio Contrato",
            "Inicio", "Final", "Días", "Salario", "Bonificación",
            "Ordinario", "Otros Ingresos", "Sueldo"
        ]
        for col, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col).value = header


    def _obtener_datos_detalle_empleado(self):
        """
        Ejecuta el SQL que genera el detalle por empleado.
        """
        company_id = self.env.user.company_id.id

        
        company_id = self.env.user.company_id.id

        sql_base = """
            SELECT r.contract_id, hc.date_start AS fecha_inicio_contrato, hc.wage AS salario_actual,
                he.id AS employee_id, he.name AS nombre_empleado,
                he.department_id,
                CONCAT(
                    COALESCE(hda.name->>'es', ''), 
                    CASE WHEN hda.name IS NULL THEN '' ELSE '/' END,
                    COALESCE(hdp.name->>'es', ''), 
                    CASE WHEN hdp.name IS NULL THEN '' ELSE '/' END,
                    COALESCE(hd.name->>'es', '')
                ) AS departamento,
                r.id, hpr.name AS nombre_nomina, hpr.date_start AS fecha_inicio_nomina,
                hpr.date_end AS fecha_final_nomina,
                COALESCE(hc.ingresos1, 0) AS ingresos1, COALESCE(hc.ingresos4, 0) AS ingresos4, 
                COALESCE(r.dias, 0) AS dias, COALESCE(hc.wage, 0) AS salario_actual,
                COALESCE(he.bonificacion_decreto, 0) AS bonificacion_decreto, COALESCE(he.bonificaciones, 0) AS bonificaciones,
                COALESCE(r.bonificacion_acumulada, 0) AS bonificacion_acumulada, COALESCE(r.sueldo_acumulado, 0) AS sueldo_acumulado, 
                COALESCE(r.salario_ordinario_anual, 0) AS salario_ordinario_anual,
                COALESCE(hc.ingresos1, 0) + COALESCE(hc.ingresos4, 0) AS bonificacion_adicional
            FROM (
                SELECT d.contract_id, d.id,
                    SUM(COALESCE(d.dias, 0)) AS dias,
                    SUM(COALESCE(d.ordinario_quincena1, 0)) + SUM(COALESCE(d.ordinario_quincena2, 0)) AS salario_ordinario_anual,
                    SUM(COALESCE(d.bonificacion, 0)) AS bonificacion_acumulada,
                    SUM(COALESCE(d.bonificacion, 0)) AS sueldo_acumulado
                FROM (
                    SELECT hp.contract_id, hpr.id,
                        SUM(CASE WHEN hsr.x_code = 'raDLA' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS dias,
                        SUM(CASE WHEN hsr.x_code = 'SO' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS salario_mensual,
                        SUM(CASE WHEN hsr.x_code = 'ORDQ1' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS ordinario_quincena1,
                        SUM(CASE WHEN hsr.x_code = 'ORDQ2' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS ordinario_quincena2,
                        SUM(CASE WHEN hsr.x_code = 'LiqAG' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS aguinaldo,
                        SUM(CASE WHEN hsr.x_code = 'LiqB14' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS bono_14,
                        SUM(CASE WHEN hsr.x_code = 'LiqB14' THEN COALESCE(hpl.total, 0) ELSE 0.0 END) AS bonificacion
                    FROM hr_payslip_run hpr
                    JOIN hr_payslip hp ON hpr.id = hp.payslip_run_id
                    JOIN hr_contract hc ON hp.contract_id = hc.id
                    JOIN res_company rc ON hp.company_id = rc.id
                    LEFT JOIN hr_payslip_line hpl ON hp.id = hpl.slip_id
                    LEFT JOIN hr_salary_rule hsr ON hsr.id = hpl.salary_rule_id
                    WHERE rc.id = %s
        """

        params = []
        if self.employee_id:
            sql_base += """
                    AND hp.employee_id = %s
                    AND hsr.x_code IN ('raDLA','SO','ORDQ1','ORDQ2','LiqAG','LiqB14')
                    AND hpr.date_start >= %s AND hpr.date_start <= %s
                    AND hpr.date_end >= hc.date_start AND hc.date_end IS NULL
                    GROUP BY hp.contract_id, hpr.id, hsr.x_code
                ) d 
                GROUP BY d.contract_id, d.id
            ) r
            JOIN hr_contract hc ON r.contract_id = hc.id
            JOIN hr_payslip_run hpr ON r.id = hpr.id
            LEFT JOIN hr_employee he ON he.id = hc.employee_id
            LEFT JOIN hr_department hd ON he.department_id = hd.id
            LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
            LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
            ORDER BY r.contract_id, r.id
            """
            params = [company_id, self.employee_id.id, self.fecha_desde, self.fecha_hasta]
        else:
            sql_base += """
                    AND hp.employee_id >= 1
                    AND hsr.x_code IN ('raDLA','SO','ORDQ1','ORDQ2','LiqAG','LiqB14')
                    AND hpr.date_start >= %s AND hpr.date_start <= %s
                    AND hpr.date_end >= hc.date_start AND hc.date_end IS NULL
                    GROUP BY hp.contract_id, hpr.id, hsr.x_code
                ) d 
                GROUP BY d.contract_id, d.id
            ) r
            JOIN hr_contract hc ON r.contract_id = hc.id
            JOIN hr_payslip_run hpr ON r.id = hpr.id
            LEFT JOIN hr_employee he ON he.id = hc.employee_id
            LEFT JOIN hr_department hd ON he.department_id = hd.id
            LEFT JOIN hr_department hdp ON hd.parent_id = hdp.id
            LEFT JOIN hr_department hda ON hdp.parent_id = hda.id
            ORDER BY r.contract_id, r.id
            """
            params = [company_id, self.fecha_desde, self.fecha_hasta]

        self.env.cr.execute(sql_base, params)
        return self.env.cr.dictfetchall()


    def _llenar_datos_detalle(self, ws, datos):
        """
        Llena la hoja del detalle con datos y totales.
        """
        from openpyxl.styles import Font

        row = 2
        wtotal_salario = 0
        wtotal_bonificacion = 0
        wtotal_ordinario = 0
        wtotal_otros_ingresos = 0
        wtotal_sueldo = 0

        for line in datos:
            ws.cell(row=row, column=1).value = line['nombre_empleado']
            ws.cell(row=row, column=2).value = line['nombre_nomina']
            ws.cell(row=row, column=3).value = line['departamento']

            if line['fecha_inicio_contrato']:
                fs = line['fecha_inicio_contrato'].strftime('%Y-%m-%d').split('-')
                ws.cell(row=row, column=4).value = f"{fs[2]}/{fs[1]}/{fs[0]}"

            if line['fecha_inicio_nomina']:
                fs = line['fecha_inicio_nomina'].strftime('%Y-%m-%d').split('-')
                ws.cell(row=row, column=5).value = f"{fs[2]}/{fs[1]}/{fs[0]}"

            if line['fecha_final_nomina']:
                fs = line['fecha_final_nomina'].strftime('%Y-%m-%d').split('-')
                ws.cell(row=row, column=6).value = f"{fs[2]}/{fs[1]}/{fs[0]}"

            ws.cell(row=row, column=7).value = round(line['dias'] or 0, 0)
            ws.cell(row=row, column=7).number_format = '#,##0'

            ws.cell(row=row, column=8).value = round(line['salario_ordinario_anual'] or 0, 2)
            ws.cell(row=row, column=8).number_format = '#,##0.00'

            ws.cell(row=row, column=9).value = round(line['bonificacion_adicional'] or 0, 2)
            ws.cell(row=row, column=9).number_format = '#,##0.00'

            ws.cell(row=row, column=10).value = round(line['bonificacion_decreto'] or 0, 2)
            ws.cell(row=row, column=10).number_format = '#,##0.00'

            ws.cell(row=row, column=11).value = round(line['bonificacion_acumulada'] or 0, 2)
            ws.cell(row=row, column=11).number_format = '#,##0.00'

            ws.cell(row=row, column=12).value = round(line['salario_ordinario_anual'] or 0, 2)
            ws.cell(row=row, column=12).number_format = '#,##0.00'

            wtotal_salario += line['salario_ordinario_anual'] or 0
            wtotal_bonificacion += line['bonificacion_adicional'] or 0
            wtotal_ordinario += line['bonificacion_acumulada'] or 0
            wtotal_otros_ingresos += line['bonificacion_acumulada'] or 0
            wtotal_sueldo += line['salario_ordinario_anual'] or 0

            row += 1

        # Totales
        ws.cell(row=row, column=1).value = 'Total: '
        ws.cell(row=row, column=1).font = Font(size=13, bold=True)

        totales = [wtotal_salario, wtotal_bonificacion, wtotal_ordinario, wtotal_otros_ingresos, wtotal_sueldo]
        for i, total in zip([8, 9, 10, 11, 12], totales):
            ws.cell(row=row, column=i).value = total
            ws.cell(row=row, column=i).number_format = '#,##0.00'
            ws.cell(row=row, column=i).font = Font(size=13, bold=True)

    def style_range(self, ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
        """
        Apply styles to a range of cells as if they were a single cell.
        :param ws:  Excel worksheet instance
        :param range: An excel range to style (e.g. A1:F20)
        :param border: An openpyxl Border
        :param fill: An openpyxl PatternFill or GradientFill
        :param font: An openpyxl Font object
        """

        top = Border(top=border.top)
        left = Border(left=border.left)
        right = Border(right=border.right)
        bottom = Border(bottom=border.bottom)

        first_cell = ws[cell_range.split(":")[0]]
        rows = list(ws[cell_range])
        for cell in rows[0]:
            cell.border = top
        for cell in rows[-1]:
            cell.border = bottom
        for row in rows:
            l = row[0]
            r = row[-1]
            l.border =  left
            r.border =  right
            if fill:
                for c in row:
                    c.fill = fill
            if font:
                for c in row:
                    c.font = font

            if alignment:
                for c in row:
                    c.alignment = alignment
        