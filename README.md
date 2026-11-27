# Gerador de Certificados — ANM (DIVCAM / SGP)

Aplicativo desktop em Python/PySide6 para gerenciamento de capacitações da **Agência Nacional de Mineração (ANM — DIVCAM / SGP)**: apuração de presenças do Microsoft Teams, aplicação de justificativas de faltas, consolidação de indicadores em Excel, emissão de certificados individuais e unificados (PowerPoint/PDF) e envio automatizado via Microsoft Outlook.

---

## Principais Recursos

- **Interface Desktop Unificada (5 etapas):**
  1. **Arquivos e Horários:** importação por arrastar-e-soltar (*drag-and-drop*) da lista de convidados e das listas de presença do Teams, com ordenação cronológica automática e edição de horário por encontro.
  2. **Dados e E-mail:** pré-preenchimento automático da carga horária (soma do *Tempo médio de participação* da Seção `1. Resumo`), datas de início/término/emissão, percentuais mínimos, lista de e-mails ignorados e editor do texto de e-mail com pré-visualização.
  3. **Relatório e Justificativas:** geração assíncrona da planilha Excel completa e matriz interativa para abonar ausências (*Justificado*) diretamente pela interface com recálculo automático.
  4. **Certificados:** conferência final de participantes aptos e geração em lote de PDFs individuais e arquivo único consolidado (`.pptx`/`.pptm` e `.pdf`) via PowerPoint COM.
  5. **Envio de E-mails:** abertura de rascunho de conferência no Outlook, disparo em massa com tabela de status em tempo real e exportação do log em `.xlsx`.
- **Recursos Globais Compartilhados:** cadastro único da lista oficial de servidores (`.xlsx`) e do modelo institucional de certificado (`.pptx`/`.pptm`) válidos para todos os cursos.
- **Armazenamento Seguro e Compatível com EDR corporativo (TrendAI Vision One):**
  - Diretório curto gerenciado automaticamente em `C:\ANM_Certificados\` (com fallback automático para `%LOCALAPPDATA%\ANM_Certificados\` se sem permissão em `C:\`).
  - Persistência versionada em JSON (`curso.json` e `app_config.json` com `schema_version = 1`).
  - Empacotamento `--onedir --windowed --noupx` sem extração de binários em `%TEMP%`.

---

## Estrutura do Projeto

```text
gerador_certificados/
│
├── app.py                      # Ponto de entrada da aplicação GUI (QT_OPENGL=software)
├── core/                       # Regras de negócio desacopladas de UI
│   ├── app_paths.py            # Gerenciamento do diretório curto C:\ANM_Certificados\
│   ├── caminhos.py             # Tratamento de caminhos longos Windows (MAX_PATH / \\?\)
│   ├── certificados.py         # Preenchimento de PPTX, conversão e junção de PDFs
│   ├── config.py               # Modelos dataclass, validações e curso.json versionado
│   ├── cursos_repo.py          # CRUD de cursos, autopreenchimento e justificativas
│   ├── emails.py               # Automação COM com o Microsoft Outlook
│   ├── eventos.py              # Contratos de log, progresso e cancelamento cooperativo
│   ├── presenca.py             # Parser Teams (1. Resumo / 2 / 3), recorte e Excel
│   └── validator.py            # Auditoria de arquivos e prevenção de caminhos profundos
│
├── gui/                        # Interface gráfica em PySide6-Essentials (QSS institucional ANM)
│   ├── file_actions.py         # Abertura no sistema e "Salvar cópia em..."
│   ├── main_window.py          # Janela principal, sidebar fixa, workers e log recolhível
│   ├── theme.py                # Paleta institucional ANM/Gov.br e QSS plano
│   ├── workers.py              # QThread com CoInitialize/CoUninitialize seguro
│   ├── tabs/                   # 5 abas do fluxo de capacitação
│   └── widgets/                # DropZoneWidget (drag-and-drop) e LogPanelWidget
│
├── tests/                      # Suíte de testes automatizados (core + GUI)
├── gerador_certificados.spec   # Configuração PyInstaller enxuta (--onedir --noupx)
├── build_executavel.py         # Script de automação do build distribuível
├── version_info.txt            # Metadados de versão Windows ANM
└── requirements.txt            # Dependências Python
```

---

## Instalação e Execução em Desenvolvimento

### Pré-requisitos
- Windows 10 ou Windows 11 (64 bits)
- Python 3.10 ou superior
- Microsoft PowerPoint e Microsoft Outlook instalados (necessários na máquina para conversão de certificados em PDF e disparo de e-mails)

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

---

## Testes Automatizados

```bash
python -m unittest discover tests
```

---

## Empacotamento do Aplicativo (`dist/GeradorCertificadosANM/`)

Gera a pasta distribuível com `GeradorCertificadosANM.exe` em modo `--onedir --windowed --noupx` usando `PySide6-Essentials` enxuto:

```bash
python build_executavel.py
```

Ao final, a pasta `dist/GeradorCertificadosANM/` conterá o executável pronto para uso ou compactação em `.zip` para distribuição (os artefatos de `build/` e `dist/` permanecem ignorados pelo `.gitignore`).

---

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
