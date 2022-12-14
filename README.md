<h1 align="center">Automatizador de Relação de Envio Posterior</h1>

<p align="center">
<img src="https://img.shields.io/badge/Automatizador-v1.8.0-%232C5263?style=for-the-badge"/>
<img src="https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54"/>
<img src="https://img.shields.io/badge/interface%20gr%C3%A1fica-tkinter-%23EE4C2C?style=for-the-badge"/>
<img src="https://img.shields.io/badge/data%20e%20hora-datetime-%23EE4C2C?style=for-the-badge"/>
<img src="https://img.shields.io/badge/excel-openpyxl-%23EE4C2C?style=for-the-badge"/>
<img src="https://img.shields.io/badge/pdf-reportlab-%23EE4C2C?style=for-the-badge"/>
<img src="https://img.shields.io/badge/gerador%20execut%C3%A1vel-pyinstaller-%2318A303?style=for-the-badge"/>
</p>

Essa aplicação foi desenvolvida, na época, para **auxiliar um colaborador** da minha antiga equipe de trabalho na **emissão de punch list** de itens para envio ao cliente.

<div align="center">
  <img src="img/interface_grafica.png"/>
</div>

Optei por hospedar o código-fonte do projeto, pois tenho um carinho especial com essa aplicação pelo fato de ter sido meu primeiro projetinho com um uso real. Além de ter sido desenvolvido no início de meus estudos :grin:

---

## :computer: Como usar?

1. Clone o repositório:
   - Copie a URL do repositório clicando em "Code", localizado na parte superior direita do repositório, no campo HTTPS;
   - Abra o Git Bash e defina o diretório em que deseja ter o código-fonte clonado;
   - Digite `git clone`, cole a URL e pressione Enter. Exemplo:
     - `git clone https://github.com/luishperna/automatizador_punch_list.git`

2. Instale as dependências (bibliotecas):
   - Abra o Git Bash e digite `pip install openpyxl` e pressione Enter para instalar a biblioteca openpyxl;
   - Após, digite `pip install reportlab` para instalar a biblioteca reportlab.
   - Caso prefira instalar todas as dependências de uma vez, digite `pip install -r requirements.txt` dentro do diretório do projeto.

3. Execute a aplicação:
   - Digite `python main.py` dentro do diretório do projeto para executar a aplicação.
   
Pronto! A aplicação já pode ser utilizada, basta **preencher os campos** e clicar em **ATUALIZAR EXCEL** (ou pressionando Enter) para gerar um Punch List novo ou atualizar o existente.

O **Punch List** ficará localizado no mesmo **diretório que o arquivo main.py**, como mostrado abaixo:

<img src="img/punch_list.png"/>

**Observação:** Não exclua e nem renomeie, tanto o arquivo **Punch list (cliente) SE.xlsx**, quanto a aba/planilha **BASE** dentro do arquivo EXCEL, isso afetará o funcionamento da aplicação.

---

## :desktop_computer: Implementação

Para fácilitar ao usuário final, foi gerado um arquivo executável (.exe) da aplicação a partir do `PyInstaller`, no qual foi responsável por juntar todos os dados em um unico arquivo .exe, sendo a linguagem de programação python e seu interpretador, bibliotecas e dependências do projeto. 

---

## :heavy_check_mark: Pré-requisitos
- [x] Python e interpretador instalados;
- [x] Gerenciador de pacotes (pip) instalado;
- [x] Git Bash instalado (OS Microsoft Windows);
- [x] Conhecimento básico em terminal.

---

## :hammer: Campos para preenchimento

Campos           | Preencher com
:-------:        | :------
SE               | O número da SE do projeto
CCM              | O número do CCM do projeto
COLUNA           | O número (mais F/T caso necessário) da coluna respectiva a pendência
GAVETA           | O número da gaveta respectiva a pendência
CÓDIGO           | A identificação/tag referente a pendência
PENDÊNCIA        | A descrição do item pendente
ESPECIFICAÇÃO    | A especificação ou detalhamento da item (campo opcional)
QUANTIDADE       | A quantidade de itens referente a pendência
RESPONSABILIDADE | O responsável por fornecer o item pendente

---

## :man_technologist: Tecnologias utilizadas

- Linguagem de programação: `python` :snake:

- Bibliotecas/módulos: `tkinter` `datetime` `openpyxl` `reportlab`

- Gerador de executável: `pyinstaller`

- Editor de Código/IDE: `visual studio code`

- Versionamento e repositório remoto: `git` `github`

---

## :x: Status do Projeto

O projeto foi **descontinuado** devido as alterações no procedimento interno da empresa, logo algumas funcionalidades podem **apresentar erros** ou **não funcionamento**, como exemplo a função de EMITIR RELATÓRIO que seria responsável por gerar um arquivo pdf com os itens preenchidos.

---

## :warning: Importante

A hospedagem/divulgação do programa e seu código-fonte foram autorizados pelo gestor do setor. Além disso, todas as informações pertinentes a empresa e os clientes foram devidamente removidos para garantir maior privacidade e segurança.

---

## Autor

| [<img src="https://avatars.githubusercontent.com/u/96630233?s=400&u=3400cfe6ba8fb87692f4f14cbdbef3e5cc996b67&v=4" width=115><br><sub>Luís Henrique Perna</sub>](https://github.com/luishperna) |
| :---: |
