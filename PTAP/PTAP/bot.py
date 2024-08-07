# Portal Transparência Aposentados e Pensionistas
from botcity.core import DesktopBot
import requirements as rq
import dados as dt
import time
import os


def not_found(label):
    print(f"Element not found: {label}")


class Bot(DesktopBot):
    caminho_completo = os.path.join(rq.diretorio_downloads, rq.nome_arquivo)
    if os.path.exists(caminho_completo):
        os.remove(caminho_completo)

    def action(self, execution=None):
        # Loga no site
        self.browse(rq.url_transparencia)
        time.sleep(15)

        # Acessa a pagina de detalhados
        if not self.find( "orgao2", matching=0.97, waiting_time=10000):
            self.not_found("orgao2")
        self.move()
        while True:
            print('while - 01')
            if self.find( "Detalhamento", matching=0.97, waiting_time=5):
                break
            else:
                self.scroll_down(40)
        self.click()
        time.sleep(5)

        # Seleciona Orgão
        self.scroll_up(1000000)
        time.sleep(2)
        if not self.find( "orgao2", matching=0.97, waiting_time=10000):
            self.not_found("orgao2")
        self.click()
        time.sleep(1)
        while True:
            print('while - 03')
            if self.find( "agencia", matching=0.97, waiting_time=5):
                break
            else:
                self.type_down()
        if not self.find( "agencia", matching=0.97, waiting_time=1000000):
            self.not_found("agencia")
        self.click()
        self.key_esc()
        time.sleep(5)

        # Exporta tabela
        if not self.find( "ponto2", matching=0.97, waiting_time=10000):
            self.not_found("ponto2")
        self.move()
        time.sleep(2)
        if not self.find( "Exportar", matching=0.97, waiting_time=10000):
            self.not_found("Exportar")
        self.click()
        # self.click_relative(0, 15)
        time.sleep(2)
        if not self.find( "ExportDados", matching=0.97, waiting_time=10000):
            self.not_found("ExportDados")
        self.click()
        
        time.sleep(2)
        if not self.find( "DadosCsv", matching=0.97, waiting_time=10000):
            self.not_found("DadosCsv")
        self.click()
        self.tab()
        self.enter()
        self.type_down()
        self.enter()
        time.sleep(2)
        self.tab()
        self.enter()
        time.sleep(15)

        # Carrega dados excel
        if dt.validador() == 'sim':
            self.alt_f4()
            dt.dados()
        else:
            print('arquivo não localizado processo de download exedeu limite de espera')


    def not_found(self, param):
        pass


if __name__ == '__main__':
    data_exe = dt.valida_exe()
    if data_exe == 'Sim':
        print('processar')
        Bot.main()
        time.sleep(20)
    else:
        print('não processar')
        time.sleep(20)
