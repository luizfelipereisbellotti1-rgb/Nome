    def carregar_capas(self):
        if not self.processo_concluido:
            messagebox.showwarning("Aviso", "Rode a atualização das capas antes de carregar o painel!")
            return

        if not os.path.exists(CAMINHO_PASTA_CAPAS):
            messagebox.showerror("Erro", f"Caminho não encontrado:\n{CAMINHO_PASTA_CAPAS}")
            return

        self.escrever_log(f"[{datetime.now().strftime('%H:%M:%S')}] Carregando dados das capas...")

        dados = []
        arquivos = [a for a in os.listdir(CAMINHO_PASTA_CAPAS) if a.endswith(".xlsx") and os.path.isfile(os.path.join(CAMINHO_PASTA_CAPAS, a))]

        if not arquivos:
            self.escrever_log("Nenhum arquivo .xlsx encontrado na pasta de capas.")
            messagebox.showwarning("Aviso", "Nenhum arquivo .xlsx encontrado na pasta de capas.")
            return

        # --- RECÁLCULO DAS FÓRMULAS NO EXCEL NATIVO (G1 e demais fórmulas) ---
        # O openpyxl não calcula fórmulas: ao salvar, ele apaga o valor em cache
        # das células com fórmula. Por isso abrimos o Excel de verdade, forçamos
        # o recálculo completo e salvamos antes de ler os valores com data_only=True.
        excel = None
        try:
            self.escrever_log("Abrindo o Excel em segundo plano para atualizar fórmulas...")
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for nome_arquivo in arquivos:
                caminho_arquivo = os.path.join(CAMINHO_PASTA_CAPAS, nome_arquivo)
                wb_excel = excel.Workbooks.Open(caminho_arquivo)
                excel.CalculateFull()
                wb_excel.Save()
                wb_excel.Close()

            excel.Quit()
            excel = None
            self.escrever_log("Fórmulas recalculadas e arquivos salvos com sucesso.")
        except Exception as e:
            self.escrever_log(f"Erro ao recalcular fórmulas no Excel: {e}")
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            messagebox.showerror("Erro", f"Erro ao recalcular fórmulas no Excel:\n{e}")
            return

        for nome_arquivo in arquivos:
            caminho_arquivo = os.path.join(CAMINHO_PASTA_CAPAS, nome_arquivo)
            match_cosif = re.match(r"^(\d+)", nome_arquivo)
            cosif_bruto = match_cosif.group(1) if match_cosif else ""
            cosif_exibicao = self._formatar_cosif(cosif_bruto) if cosif_bruto else nome_arquivo

            try:
                wb = openpyxl.load_workbook(caminho_arquivo, data_only=True)
                if "CONCILIAÇÃO" not in wb.sheetnames:
                    self.escrever_log(f"Aviso: Aba 'CONCILIAÇÃO' não encontrada em '{nome_arquivo}'.")
                    wb.close()
                    continue

                ws = wb["CONCILIAÇÃO"]
                descricao = ws["B1"].value or ""
                saldo_inicial = self._valor_numerico(ws["E1"].value)
                debito = self._valor_numerico(ws["E2"].value)
                credito = self._valor_numerico(ws["E3"].value)
                saldo_final = self._valor_numerico(ws["E4"].value)
                open_outros = self._valor_numerico(ws["G1"].value)
                conciliacao = saldo_final - open_outros

                dados.append({
                    "cosif": cosif_exibicao,
                    "descricao": str(descricao),
                    "saldo_inicial": saldo_inicial,
                    "debito": debito,
                    "credito": credito,
                    "saldo_final": saldo_final,
                    "open_outros": open_outros,
                    "conciliacao": conciliacao,
                })
                wb.close()
            except Exception as e:
                self.escrever_log(f"Erro ao ler '{nome_arquivo}': {e}")

        dados.sort(key=lambda d: d["cosif"])
        self.dados_capas = dados
        self.montar_painel_capas(dados)
        self.escrever_log(f"[{datetime.now().strftime('%H:%M:%S')}] Painel carregado com {len(dados)} capas.")
