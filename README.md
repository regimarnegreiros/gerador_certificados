# Gerador de Certificados

Sistema automatizado para gerenciamento de cursos, validação de listas de presença do Microsoft Teams, consolidação de dados em planilhas Excel, geração de certificados individuais e unificados (PowerPoint/PDF) e envio automatizado via Microsoft Outlook.

Desenvolvido para a **Agência Nacional de Mineração (ANM)** para automatizar os fluxos de capacitação e gestão de pessoas.

---

## 📁 Estrutura do Projeto

```text
gerador_certificados/
│
├── core/                       # Módulos centrais reutilizáveis
│   ├── caminhos.py             # Tratamento de caminhos longos do Windows (MAX_PATH / \\?\)
│   ├── certificados.py         # Manipulação de PPTX e conversão/junção de PDFs
│   ├── config.py               # Leitura e parsing de variaveis.txt
│   ├── emails.py               # Integração COM com o Microsoft Outlook
│   ├── presenca.py             # Processamento de presenças, cruzamento de dados e Excel
│   ├── ui.py                   # Interface de terminal (Cores ANSI, menus e prompts)
│   └── validator.py            # Validação estrutural de arquivos e variáveis
│
├── templates/                  # Modelos padrão copiados para cada curso criado
│   ├── leia-me_template.txt    # Instruções para o usuário final
│   └── variaveis_template.txt  # Arquivo de configuração de cada curso
│
├── tests/                      # Suíte de testes automatizados
│   ├── test_suite.py           # Testes unitários de todos os módulos
│   └── test_e2e_pipeline.py    # Teste de ponta a ponta (E2E) com dados sintéticos
│
├── criar_curso.py              # Script/Executável principal de gerenciamento
├── gerar_relatorio.py          # Etapa 1: Apuração de presenças e geração do Excel
├── gerar_certificados.py       # Etapa 2: Geração dos certificados (PPTX/PDF)
├── enviar_emails.py            # Etapa 3: Disparo de e-mails com certificados anexos
│
├── build_executaveis.py        # Automação de compilação dos 4 binários via PyInstaller
├── version_info.txt            # Metadados de versão Windows para os executáveis
├── requirements.txt            # Dependências Python do projeto
└── .gitignore                  # Regras de exclusão do Git
```

---

## 🚀 Instalação e Ambiente de Desenvolvimento

### Pré-requisitos
- Python 3.10 ou superior (testado no Windows)
- Microsoft PowerPoint e Microsoft Outlook instalados (necessários para a conversão de PPTX em PDF e envio de e-mails)

### 1. Clonar o repositório
```bash
git clone https://github.com/SEU_USUARIO/gerador_certificados.git
cd gerador_certificados
```

### 2. Criar e ativar um ambiente virtual (recomendado)
```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Instalar as dependências
```bash
pip install -r requirements.txt
```

---

## 🧪 Testes Automatizados

Para executar toda a suíte de testes unitários e de integração:

```bash
python -m unittest discover tests
```

---

## 📦 Compilação dos Executáveis (.exe)

O script `build_executaveis.py` automatiza a compilação de todos os 4 executáveis com PyInstaller em modo `--onefile`:

```bash
# Compilar todos os 4 executáveis:
python build_executaveis.py

# Ou forçar recompilação completa:
python build_executaveis.py --force
```

> **Nota sobre Releases:** Os executáveis gerados (`.exe`) são ignorados pelo Git devido ao tamanho dos binários. Para distribuí-los aos usuários finais, utilize a aba de **Releases** do GitHub.

---

## 📄 Licença

Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` para mais detalhes.

