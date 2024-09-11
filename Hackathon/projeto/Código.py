import pandas as pd
import googlemaps
import requests
import os
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl import Workbook

# Carrega as variáveis de ambiente do arquivo .env-------------------------------------------
load_dotenv()

# Obtém a chave da API do Google Maps.-----------------------------------------------------------
API_KEY = os.getenv('GOOGLE_MAPS_API_KEY')
gmaps = googlemaps.Client(key=API_KEY)

# Função para calculo da distância e do custo da viagem usando a API do Google Maps----------------
def calcular_distancia_e_custo(origem, destino, chave_api, taxa_por_km=0.50):
    url = "https://maps.googleapis.com/maps/api/distancematrix/json"
    parametros = {
        "origins": origem,
        "destinations": destino,
        "key": chave_api,
        "mode": "driving",  
        "language": "pt-BR",  
        "units": "metric"  
    }

    resposta = requests.get(url, params=parametros)
    dados = resposta.json()

    if dados["status"] == "OK":
        distancia_metros = dados["rows"][0]["elements"][0]["distance"]["value"]
        distancia_km = distancia_metros / 1000  
        custo = distancia_km * taxa_por_km  
        return distancia_km, custo
    else:
        raise Exception(f"Erro na API do Google Maps: {dados['error_message']}")

# Definição da classe Veículo
class Veiculo:
    def __init__(self, modelo, marca, ano, placa, tipo, categoria, status='disponível'):
        self.modelo = modelo
        self.marca = marca
        self.ano = ano
        self.placa = placa
        self.tipo = tipo
        self.categoria = categoria
        self.status = status

    def __repr__(self):
        return f"{self.modelo} ({self.placa}) - {self.tipo} - {self.categoria} - {self.status}"

# Definição da classe cliente
class Cliente:
    def __init__(self, nome, cpf, telefone, email):
        self.nome = nome
        self.cpf = cpf
        self.telefone = telefone
        self.email = email
        self.pontos_fidelidade = 0  # Inicializa pontos de fidelidade com zero

    def adicionar_pontos(self, pontos):
        self.pontos_fidelidade += pontos
        print(f"{pontos} pontos adicionados. Total de pontos: {self.pontos_fidelidade}")

    def usar_pontos(self, pontos):
        if pontos <= self.pontos_fidelidade:
            self.pontos_fidelidade -= pontos
            desconto = pontos * 0.10  # Cada ponto vale 10 centavos
            print(f"{pontos} pontos usados para desconto de R${desconto:.2f}.")
            return desconto
        else:
            print("Pontos insuficientes.")
            return 0
# Definição da classe locação
class Locacao:
    CATEGORIA_PRECO = {
        'Ferro': 100,
        'Ouro': 150,
        'Premium': 200
    }
    
    def __init__(self, cliente, veiculo, data_retirada, data_devolucao_prevista, distancia_km=0, taxa_por_km=0.50):
        self.cliente = cliente
        self.veiculo = veiculo
        self.data_retirada = data_retirada
        self.data_devolucao_prevista = data_devolucao_prevista
        self.data_devolucao_real = None  # Inicializa a data de devolução real como None
        self.distancia_km = distancia_km  # Armazena a distância percorrida
        self.taxa_por_km = taxa_por_km  # Armazena a taxa por km para multa

    def calcular_preco_total(self, multa=0):
        # Calcula os dias de aluguel
        dias_alugados = (self.data_devolucao_prevista - self.data_retirada).days + 1

        # Obtém o preço base por dia baseado na categoria do veículo
        preco_base_diario = Locacao.CATEGORIA_PRECO.get(self.veiculo.categoria, 100)  # Preço base por dia
        preco_base = preco_base_diario * dias_alugados
        
        # Adiciona a multa ao preço base
        total_a_pagar = preco_base + multa
        return preco_base, total_a_pagar

    def calcular_multa(self):
        if self.data_devolucao_real and self.data_devolucao_real > self.data_devolucao_prevista:
            multa = self.distancia_km * self.taxa_por_km  # Multa é igual à distância percorrida vezes a taxa por km
            return multa
        return 0

    def acumular_pontos(self):
        preco_base, total_a_pagar = self.calcular_preco_total()
        pontos = int(total_a_pagar // 10)  # 1 ponto para cada R$10 gastos
        self.cliente.adicionar_pontos(pontos)

# Definição da classe histórico de locações
class HistoricoLocacao:
    def __init__(self):
        self.locacoes = []

    def adicionar_locacao(self, locacao):
        self.locacoes.append(locacao)

    def listar_historico(self):
        return self.locacoes

# Inicializa o histórico de locações
historico = HistoricoLocacao()

# Função para buscar veículos disponíveis com base no tipo e status
def buscar_veiculos(tipo=None, status=None, categoria=None):
    veiculos_filtrados = veiculos
    
    if tipo:
        veiculos_filtrados = [v for v in veiculos_filtrados if v.tipo.lower() == tipo.lower()]
    if status:
        veiculos_filtrados = [v for v in veiculos_filtrados if v.status.lower() == status.lower()]
    if categoria:
        veiculos_filtrados = [v for v in veiculos_filtrados if v.categoria.lower() == categoria.lower()]

    if veiculos_filtrados:
        df_veiculos_filtrados = pd.DataFrame([vars(v) for v in veiculos_filtrados])
        print("\nResultados da Busca de Veículos:")
        print(df_veiculos_filtrados.to_string(index=False))
    else:
        print("Nenhum veículo encontrado com os filtros fornecidos.")
        
# Função para adicionar um veículo à lista de veículos
def adicionar_veiculo(veiculo):
    veiculos.append(veiculo)
    salvar_dados()
# Função para buscar locações com base em filtros
def buscar_locacoes(nome_cliente=None, cpf_cliente=None, placa_veiculo=None, tipo_veiculo=None, categoria_veiculo=None, data_inicio=None, data_fim=None):
    locacoes_filtradas = historico.listar_historico()

    if nome_cliente:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.cliente.nome.lower() == nome_cliente.lower()]
    if cpf_cliente:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.cliente.cpf == cpf_cliente]
    if placa_veiculo:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.veiculo.placa == placa_veiculo]
    if tipo_veiculo:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.veiculo.tipo.lower() == tipo_veiculo.lower()]
    if categoria_veiculo:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.veiculo.categoria.lower() == categoria_veiculo.lower()]
    if data_inicio:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.data_retirada >= data_inicio]
    if data_fim:
        locacoes_filtradas = [l for l in locacoes_filtradas if l.data_devolucao_prevista <= data_fim]

    if locacoes_filtradas:
        df_locacoes_filtradas = pd.DataFrame([{
            'Cliente': l.cliente.nome,
            'Veículo': l.veiculo.modelo,
            'Placa': l.veiculo.placa,
            'Data Retirada': l.data_retirada,
            'Data Devolução Prevista': l.data_devolucao_prevista,
            'Data Devolução Real': l.data_devolucao_real,
            'Distância (km)': l.distancia_km
        } for l in locacoes_filtradas])
        print("\nResultados da Busca de Locações:")
        print(df_locacoes_filtradas.to_string(index=False))
    else:
        print("Nenhuma locação encontrada com os filtros fornecidos.")


# Função para listar todos os veículos cadastrados
def listar_veiculos():
    df_veiculos = pd.DataFrame([vars(v) for v in veiculos])
    print("\nListagem de Veículos:")
    print(df_veiculos.to_string(index=False))

# Função para alugar um veículo
def alugar_veiculo(cliente, veiculo, data_retirada, data_devolucao_prevista, origem, destino, chave_api):
    if veiculo.status == 'disponível':
        try:
            distancia, _ = calcular_distancia_e_custo(origem, destino, chave_api, taxa_por_km=0.50)
        except Exception as e:
            print(f"Erro ao calcular a distância: {e}")
            return None
        
        locacao = Locacao(cliente, veiculo, data_retirada, data_devolucao_prevista, distancia_km=distancia, taxa_por_km=0.50)
        
        # Atualiza o status do veículo antes de adicionar ao histórico e salvar os dados
        veiculo.status = 'alugado'
        
        # Adiciona a locação ao histórico e salva os dados
        historico.adicionar_locacao(locacao)
        salvar_dados()
        
        # Acumula pontos após a locação
        locacao.acumular_pontos()
        
        # Exibe informações sobre a locação
        print(f"Locação realizada com sucesso!")
        print(f"Origem: {origem}")
        print(f"Destino: {destino}")
        print(f"Distância: {distancia:.2f} km")
        print(f"Multa, se houver atraso na entrega do veículo, com base na distância de {distancia:.2f} km e taxa de R$ 0,50 por km.")
        
        return locacao
    else:
        raise ValueError("Veículo não disponível")

# Função para devolver um veículo
def devolver_veiculo(locacao, data_devolucao_real):
    locacao.data_devolucao_real = data_devolucao_real
    locacao.veiculo.status = 'disponível'
    multa = locacao.calcular_multa()
    preco_base, total_a_pagar = locacao.calcular_preco_total(multa)
    
    print(f"Veículo {locacao.veiculo.placa} devolvido com sucesso!")
    print(f"Preço base (sem multa): R${preco_base:.2f}")
    
    if multa > 0:
        print(f"Multa por atraso (baseada na distância de {locacao.distancia_km:.2f} km): R${multa:.2f}")
    
    # Pergunta ao usuário se deseja usar os pontos de fidelidade
    usar_pontos = input(f"Você deseja usar seus pontos de fidelidade para um desconto? (Você tem {locacao.cliente.pontos_fidelidade} pontos disponíveis) [s/n]: ").strip().lower()
    
    if usar_pontos == 's':
        pontos_usados = int(input("Quantos pontos você deseja usar? "))
        desconto = locacao.cliente.usar_pontos(pontos_usados)
        total_a_pagar -= desconto
        print(f"Desconto aplicado: R${desconto:.2f}")
    else:
        print("Nenhum desconto aplicado.")
    
    print(f"Total a pagar após desconto: R${total_a_pagar:.2f}")
    salvar_dados()

# Função para carregar os dados de um arquivo Excel
def carregar_dados():
    global veiculos, clientes, historico
    try:
        # Verifica se o arquivo existe
        if not os.path.exists('banco de dados.xlsx'):
            # Se não existir, cria um novo arquivo Excel
            workbook = Workbook()
            workbook.save('banco de dados.xlsx')

        # Carrega os dados dos veículos
        df_veiculos = pd.read_excel('banco de dados.xlsx', sheet_name='Veiculos', engine='openpyxl')
        veiculos = [Veiculo(row['modelo'], row['marca'], row['ano'], row['placa'], row['tipo'],
                            row['categoria'], row.get('status', 'disponível')) for _, row in df_veiculos.iterrows()]

        # Carrega os dados dos clientes
        df_clientes = pd.read_excel('banco de dados.xlsx', sheet_name='Clientes', engine='openpyxl')
        clientes = [Cliente(row['nome'], row['cpf'], row['telefone'], row['email'])
                    for _, row in df_clientes.iterrows()]

        # Carrega os dados das locações
        df_locacoes = pd.read_excel('banco de dados.xlsx', sheet_name='Locacoes', engine='openpyxl')
        for _, row in df_locacoes.iterrows():
            cliente = next((c for c in clientes if c.cpf == row['cpf']), None)
            veiculo = next((v for v in veiculos if v.placa == row['placa']), None)

            if cliente is None or veiculo is None:
                continue

            data_retirada = pd.to_datetime(row['data_retirada'])
            data_devolucao_prevista = pd.to_datetime(row['data_devolucao_prevista'])
            distancia_km = row.get('distancia_km', 0)
            locacao = Locacao(cliente, veiculo, data_retirada, data_devolucao_prevista, distancia_km=distancia_km)
            locacao.data_devolucao_real = pd.to_datetime(row['data_devolucao_real']) if not pd.isna(row['data_devolucao_real']) else None
            historico.adicionar_locacao(locacao)
    except Exception as e:
        print(f"Erro ao carregar dados: {e}")

# Função para salvar os dados em um arquivo Excel
def salvar_dados():
    global veiculos, clientes, historico
    try:
        # Carrega o workbook existente
        with pd.ExcelWriter('banco de dados.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # Salva os dados dos veículos
            df_veiculos = pd.DataFrame([vars(v) for v in veiculos])
            df_veiculos.to_excel(writer, sheet_name='Veiculos', index=False)

            # Salva os dados dos clientes
            df_clientes = pd.DataFrame([vars(c) for c in clientes])
            df_clientes.to_excel(writer, sheet_name='Clientes', index=False)

            # Salva os dados das locações
            locacoes_data = [{
                'nome': l.cliente.nome,
                'cpf': l.cliente.cpf,
                'placa': l.veiculo.placa,
                'data_retirada': l.data_retirada,
                'data_devolucao_prevista': l.data_devolucao_prevista,
                'data_devolucao_real': l.data_devolucao_real,
                'distancia_km': l.distancia_km
            } for l in historico.listar_historico()]
            df_locacoes = pd.DataFrame(locacoes_data)
            df_locacoes.to_excel(writer, sheet_name='Locacoes', index=False)

    except Exception as e:
        print(f"Erro ao salvar dados: {e}")

# Função para exibir o menu principal
def menu():
    carregar_dados()
    while True:
        print("\nMenu:")
        print("1. Adicionar Veículo")
        print("2. Listar Veículos")
        print("3. Marcar Veículo em Manutenção")
        print("4. Concluir Manutenção")
        print("5. Adicionar Cliente")
        print("6. Listar Clientes")
        print("7. Alugar Veículo")
        print("8. Devolver Veículo")
        print("9. Relatório de Locações")
        print("10. Relatório de Quantidade de Veículos")
        print("11. Buscar Locações")
        print("12. Sair")
        escolha = input("Escolha uma opção: ")

        if escolha == "1":
            modelo = input("Modelo: ")
            marca = input("Marca: ")
            ano = input("Ano: ")
            placa = input("Placa: ")
            tipo = input("Tipo (carro/moto): ")
            categoria = input("Categoria (Ferro/Ouro/Premium): ")
            veiculo = Veiculo(modelo, marca, ano, placa, tipo, categoria)
            adicionar_veiculo(veiculo)
            print(f"Veículo {placa} cadastrado com sucesso.")
        elif escolha == "2":
            listar_veiculos()
        elif escolha == "3":
            placa = input("Placa do veículo a ser marcado em manutenção: ")
            veiculo = next((v for v in veiculos if v.placa == placa), None)
            if veiculo:
                veiculo.status = 'em manutenção'
                salvar_dados()
                print(f"Veículo {placa} marcado como em manutenção.")
            else:
                print("Veículo não encontrado.")
        elif escolha == "4":
            placa = input("Placa do veículo a ser retirado da manutenção: ")
            veiculo = next((v for v in veiculos if v.placa == placa), None)
            if veiculo:
                veiculo.status = 'disponível'
                salvar_dados()
                print(f"Veículo {placa} concluído de manutenção e está disponível.")
            else:
                print("Veículo não encontrado.")
        elif escolha == "5":
            nome = input("Nome: ")
            cpf = input("CPF: ")
            telefone = input("Telefone: ")
            email = input("Email: ")
            cliente = Cliente(nome, cpf, telefone, email)
            clientes.append(cliente)
            salvar_dados()
            print(f"Cliente {nome} adicionado com sucesso!")
        elif escolha == "6":
            df_clientes = pd.DataFrame([vars(c) for c in clientes])
            print("\nListagem de Clientes:")
            print(df_clientes.to_string(index=False))
        elif escolha == "7":
            cpf = input("CPF do cliente: ")
            cliente = next((c for c in clientes if c.cpf == cpf), None)
            if cliente:
                origem = input("Qual é a origem da viagem? ")
                destino = input("Qual é o destino da sua viagem? ")
                data_retirada = pd.to_datetime(input("Data de Retirada (YYYY-MM-DD): "))
                data_devolucao_prevista = pd.to_datetime(input("Data de Devolução Prevista (YYYY-MM-DD): "))
                tipo = input("Filtrar por Tipo (carro/moto) [pressione Enter para todos]: ")
                categoria = input("Filtrar por Categoria (Ferro/Ouro/Premium) [pressione Enter para todos]: ")
                buscar_veiculos(tipo=tipo if tipo else None, status='disponível', categoria=categoria if categoria else None)

                placa = input("Placa do veículo que deseja alugar: ")
                veiculo = next((v for v in veiculos if v.placa == placa), None)
                if veiculo:
                    try:
                        locacao = alugar_veiculo(cliente, veiculo, data_retirada, data_devolucao_prevista, origem, destino, API_KEY)
                        print(f"Veículo {veiculo.placa} alugado com sucesso!")
                    except ValueError as e:
                        print(e)
                else:
                    print("Veículo não encontrado.")
            else:
                print("Cliente não encontrado.")
        elif escolha == "8":
            cpf = input("CPF do cliente: ")
            placa = input("Placa do veículo: ")
            data_devolucao_real = pd.to_datetime(input("Data de Devolução Real (YYYY-MM-DD): "))
            locacao = next((l for l in historico.listar_historico()
                            if l.cliente.cpf == cpf and l.veiculo.placa == placa and l.data_devolucao_real is None), None)
            if locacao:
                devolver_veiculo(locacao, data_devolucao_real)
            else:
                print("Locação não encontrada ou já devolvida.")
        elif escolha == "9":
            print("\nRelatório de Locações:")
            locacoes = historico.listar_historico()
            df_locacoes = pd.DataFrame([{
                'Cliente': l.cliente.nome,
                'Veículo': l.veiculo.modelo,
                'Data Retirada': l.data_retirada,
                'Data Devolução Prevista': l.data_devolucao_prevista,
                'Data Devolução Real': l.data_devolucao_real
            } for l in locacoes])
            print(df_locacoes.to_string(index=False))
        elif escolha == "10":
            qtd_alugados = sum(v.status == 'alugado' for v in veiculos)
            qtd_disponiveis = sum(v.status == 'disponível' for v in veiculos)
            qtd_manutencao = sum(v.status == 'em manutenção' for v in veiculos)
            df_veiculos = pd.DataFrame({
                'Status': ['Alugados', 'Disponíveis', 'Em Manutenção'],
                'Quantidade': [qtd_alugados, qtd_disponiveis, qtd_manutencao]
            })
            print("\nRelatório de Veículos:")
            print(df_veiculos.to_string(index=False))
        elif escolha == "11":
            nome_cliente = input("Nome do Cliente [pressione Enter para todos]: ")
            cpf_cliente = input("CPF do Cliente [pressione Enter para todos]: ")
            placa_veiculo = input("Placa do Veículo [pressione Enter para todos]: ")
            tipo_veiculo = input("Tipo do Veículo [pressione Enter para todos]: ")
            categoria_veiculo = input("Categoria do Veículo [pressione Enter para todos]: ")
            data_inicio = input("Data Início (YYYY-MM-DD) [pressione Enter para todos]: ")
            data_fim = input("Data Fim (YYYY-MM-DD) [pressione Enter para todos]: ")

            data_inicio = pd.to_datetime(data_inicio) if data_inicio else None
            data_fim = pd.to_datetime(data_fim) if data_fim else None

            buscar_locacoes(
                nome_cliente=nome_cliente if nome_cliente else None,
                cpf_cliente=cpf_cliente if cpf_cliente else None,
                placa_veiculo=placa_veiculo if placa_veiculo else None,
                tipo_veiculo=tipo_veiculo if tipo_veiculo else None,
                categoria_veiculo=categoria_veiculo if categoria_veiculo else None,
                data_inicio=data_inicio,
                data_fim=data_fim
            )
        elif escolha == "12":
            salvar_dados()
            print("Dados salvos e programa encerrado.")
            break
        else:
            print("Opção inválida. Tente novamente.")

# Executa o menu
if __name__ == "__main__":
    # Inicializa as listas globais de veículos e clientes
    veiculos = []
    clientes = []
    menu()