def atualizar_status_cliente(cpf_cliente, novo_status):
    """
    Lê o CSV, atualiza o status de um cliente específico pelo CPF e salva o arquivo.
    """
    try:
        # Lê o arquivo CSV mais recente para evitar sobrescrever dados antigos
        df = pd.read_csv('base_clientes_fake.csv', sep=';')
        
        # Encontra a linha correspondente ao CPF
        if cpf_cliente in df['cpf'].astype(str).values:
            df.loc[df['cpf'].astype(str) == cpf_cliente, 'status'] = novo_status
            
            # Salva o DataFrame inteiro de volta no arquivo CSV
            df.to_csv('base_clientes_fake.csv', sep=';', index=False)
            print(f"✅ Status do CPF {cpf_cliente} atualizado para '{novo_status}' no arquivo 'base_clientes_fake.csv'")
        else:
            print(f"⚠️ CPF {cpf_cliente} não encontrado no arquivo CSV. Nenhuma atualização foi feita.")
            
    except FileNotFoundError:
        print("❌ Erro: Arquivo 'base_clientes_fake.csv' não encontrado.")
    except Exception as e:
        print(f"❌ Ocorreu um erro ao atualizar o status no CSV: {e}")



#=======================================================================================================================
#                  FUNÇÃO 2 - PRINCIPAL - Buscar consórcio do cliente e selecionar a melhor opção
#=======================================================================================================================

def buscar_consorcio_cliente():
    global grupo_encontrado, cpf_atual, cliente_atual
    global driver
    
    # ... (o início da sua função continua exatamente igual) ...
    # ... (carregamento da lista de grupos, validação do CPF, extração do valor máximo) ...
    # ... (loop pelas linhas da tabela de grupos) ...

    # ===================================================
    # DENTRO DO SEU LOOP `for linha in linhas_creditos:`
    # QUANDO A MELHOR OPÇÃO É ENCONTRADA E CLICADA
    # ===================================================
    
    # Substitua o final do seu loop de grupos por este trecho:
    
                        # ... (código para clicar no crédito, retirar seguro, etc.) ...
                        
                        ### >>> Clicar em CONTRATAR COTA
                        botao_contratar_cota = driver.find_element(By.XPATH, '//span[contains(text(), " contratar cota ")]')
                        action.move_to_element(botao_contratar_cota).pause(random.uniform(0.2, 0.7)).click(botao_contratar_cota).perform()
                        human_sleep(1, 2)
                        print("Clicado em CONTRATAR COTA, aguardando próxima tela...")

                        # Chama a função para preencher os dados pessoais.
                        # A função retornará True se for bem-sucedida.
                        sucesso_preenchimento = preencher_dados_pessoais()

                        # SE o preenchimento e o clique final foram bem-sucedidos
                        if sucesso_preenchimento:
                            # Chama a nova função para ATUALIZAR O STATUS NO ARQUIVO
                            atualizar_status_cliente(cpf_atual, 'Concluido')
                        else:
                            # Se falhou, atualiza como 'Erro' para revisão manual
                            print(f"⚠️ O preenchimento para o CPF {cpf_atual} falhou. Marcando como 'Erro'.")
                            atualizar_status_cliente(cpf_atual, 'Erro')

                        # Sai do loop de créditos, pois o trabalho foi feito
                        break 
                
            # ... (resto da lógica de voltar ou sair do loop de grupos) ...
