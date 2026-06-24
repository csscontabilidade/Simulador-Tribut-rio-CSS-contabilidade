    def gerar(self):
        import unicodedata

        def limpar(texto):
            # Remove acentos e caracteres especiais, mantendo apenas ASCII compatível com latin-1
            if not isinstance(texto, str):
                texto = str(texto)
            nfkd = unicodedata.normalize('NFKD', texto)
            return nfkd.encode('ascii', 'ignore').decode('ascii')

        pdf = FPDF()
        pdf.add_page()
        pdf.set_fill_color(26, 60, 94)
        pdf.rect(0, 0, 210, 35, 'F')
        pdf.set_y(8)
        pdf.set_font("Arial", "B", 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 8, limpar(ESCRITORIO['nome']), ln=True, align="C")
        pdf.set_font("Arial", "", 9)
        pdf.cell(0, 6, limpar(f"Contador: {ESCRITORIO['contador']} | {ESCRITORIO['crc']}"), ln=True, align="C")
        pdf.set_y(40)
        pdf.set_text_color(0, 0, 0)

        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, limpar("1. DADOS DA EMPRESA"), ln=True)
        pdf.set_font("Arial", "", 10)
        info = [
            limpar(f"Nome: {self.emp.nome}"),
            limpar(f"Regime Atual: {self.emp.regime_atual}"),
            limpar(f"Faturamento Mensal: R$ {self.emp.faturamento_mensal:,.2f}   Anual: R$ {self.emp.fat_anual:,.2f}"),
            limpar(f"Folha Mensal: R$ {self.emp.folha_mensal:,.2f}   Anual (13º): R$ {self.emp.folha_anual:,.2f}")
        ]
        for i in info:
            pdf.cell(0, 6, i, ln=True)
        pdf.ln(4)

        # 2026
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, limpar("2. CENÁRIO ATUAL (2026)"), ln=True)
        for r in self.res26.values():
            pdf.set_font("Arial", "B", 10)
            pdf.cell(0, 7, limpar(r['regime']), ln=True)
            pdf.set_font("Arial", "", 9)
            pdf.cell(0, 6, limpar(f"Carga Total: R$ {r['carga_total']:,.2f}   Alíquota efetiva: {r['aliquota_efetiva']:.2f}%"), ln=True)
            for k, v in r['detalhes'].items():
                pdf.cell(0, 6, limpar(f"  {k}: R$ {v:,.2f}"), ln=True)
            pdf.ln(2)

        # 2027
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, limpar("3. CENÁRIO 2027 (COM CBS)"), ln=True)
        for r in self.res27.values():
            pdf.set_font("Arial", "B", 10)
            pdf.cell(0, 7, limpar(r['regime']), ln=True)
            pdf.set_font("Arial", "", 9)
            pdf.cell(0, 6, limpar(f"Carga Total: R$ {r['carga_total']:,.2f}   Alíquota efetiva: {r['aliquota_efetiva']:.2f}%"), ln=True)
            for k, v in r['detalhes'].items():
                pdf.cell(0, 6, limpar(f"  {k}: R$ {v:,.2f}"), ln=True)
            pdf.ln(2)

        # Split
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, limpar("4. SPLIT PAYMENT (CBS)"), ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 6, limpar(f"CBS retida mensal: R$ {self.sp['split']:,.2f}"), ln=True)
        pdf.cell(0, 6, limpar(f"Recebimento líquido: R$ {self.sp['receb_liq']:,.2f}"), ln=True)
        pdf.cell(0, 6, limpar(f"Capital de giro sugerido: R$ {self.sp['nec_liquida']:,.2f}"), ln=True)

        # Formação e experiência
        pdf.add_page()
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, limpar("5. CONSULTOR RESPONSÁVEL"), ln=True)
        pdf.set_font("Arial", "", 10)
        pdf.cell(0, 6, limpar(f"{ESCRITORIO['contador']} – {ESCRITORIO['crc']}"), ln=True)
        pdf.ln(3)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 6, limpar("Formação:"), ln=True)
        pdf.set_font("Arial", "", 9)
        for f in ESCRITORIO['formacao']:
            pdf.cell(0, 6, limpar(f"• {f}"), ln=True)
        pdf.ln(2)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(0, 6, limpar("Experiência:"), ln=True)
        for e in ESCRITORIO['experiencia']:
            pdf.cell(0, 6, limpar(f"• {e}"), ln=True)

        # Retornar bytes diretamente (não usar .encode)
        return pdf.output(dest='S')