# 🛡️ Forensic Incident Response Framework (DFIR)

Um framework automatizado em PowerShell desenvolvido para **Triagem Forense Computacional**, **Resposta a Incidentes (IR)** e **Auditoria de Conformidade em Segurança da Informação**. 

O objetivo principal desta suíte é permitir a busca profunda e automatizada por Indicadores de Comprometimento (IoCs), vazamento de dados sensíveis e vestígios de artefatos maliciosos (como extensões de ransomware) em sistemas de arquivos Windows, gerando evidências criptografadas e relatórios gerenciais prontos para auditoria.

---

## 🌟 Principais Recursos

- 🔍 **Varredura Profunda Multi-formato**: Suporte nativo para documentos do Microsoft Office (`.docx`, `.xlsx`, `.pptx`), PDFs, logs e arquivos de texto simples.
- 🚨 **Detecção de Ransomware**: Identificação automatizada de extensões críticas conhecidas (ex: `.crypt`, `.locked`).
- 🔐 **Quarentena Segura & Criptografada**: Isolamento automático de arquivos suspeitos com criação de pacote ZIP protegido por senha via 7-Zip (`-mhe=on` com criptografia de cabeçalho).
- ⛓️ **Cadeia de Custódia (Integridade Hash)**: Geração de hashes **SHA256** para cada artefato identificado, garantindo a rastreabilidade e não repúdio da evidência.
- 📊 **Relatório Executivo HTML**: Emissão de relatórios estruturados com métricas de tempo de execução, termos buscados, perito responsável e dados detalhados.
- 🧹 **Gestão de Recursos & Processos**: Módulo integrado para encerramento limpo de processos "zumbis" do Office (`EXCEL`, `WINWORD`, `POWERPNT`).

---

## 📁 Estrutura do Repositório

```text
forensic-incident-response/
├── .gitignore
├── README.md
└── src/
    ├── .gitignore
    ├── Invoke-ForensicAudit.ps1   # Engine principal de Auditoria Forense
    └── Clear-OfficeProcesses.ps1  # Utilitário de limpeza de processos COM
```

---

## 🚀 Como Usar

### Pré-requisitos
- **PowerShell 5.1** ou superior.
- **Privilégios de Administrador** (necessário para acesso a diretórios de sistema e instanciação de objetos COM).
- **Microsoft Office** instalado na máquina auditada.
- **7-Zip** instalado no caminho padrão (`C:\Program Files\7-Zip\7z.exe`) — *Necessário para quarentena criptografada*.

---

### 1. Executando uma Auditoria Forense Simples

Varredura recursiva no diretório especificado buscando por termos específicos ou extensões suspeitas:

```powershell
.\src\Invoke-ForensicAudit.ps1 -Termos "senha", "confidencial", ".crypt" -CaminhoAlvo "D:\Documentos" -RelatorioNome "Relatorio_Auditoria.html"
```

### 2. Executando Auditoria Completa com Quarentena Ativada

Realiza a varredura, isola os arquivos encontrados no Desktop, calcula os hashes SHA256 e compacta a quarentena em um arquivo ZIP protegido por senha:

```powershell
.\src\Invoke-ForensicAudit.ps1 -Termos "GRISELI", "DALLAGNOL" -CaminhoAlvo "C:\Users\Public" -AtivarQuarentena -SenhaZip "SenhaForteEvidencia123"
```

### 3. Limpeza de Instâncias "Zumbis" do Office

Após execuções intensivas ou em caso de interrupção abrupta de rotinas COM, limpe instâncias ocultas que ficaram presas na memória:

```powershell
.\src\Clear-OfficeProcesses.ps1
```

---

## 📊 Exemplo do Relatório Gerado (HTML)

O script gera um relatório HTML moderno com visual escuro/corporativo contendo:
- **Metadados:** Perito responsável, tempo total de execução e termos pesquisados.
- **Tabela de Evidências:** Data/Hora, Nome do Arquivo, Tipo de Incidente, Ação Executada, Caminho Completo e Hash SHA256.

---

## ⚙️ Parâmetros do Script Principal

| Parâmetro | Tipo | Obrigatório | Descrição |
| :--- | :---: | :---: | :--- |
| `-Termos` | `String[]` | **Sim** | Lista de strings ou extensões a serem pesquisadas no conteúdo ou nome do arquivo. |
| `-CaminhoAlvo` | `String` | Não | Diretório inicial da varredura. Padrão: pasta atual (`.`). |
| `-RelatorioNome`| `String` | Não | Nome do arquivo HTML de saída. Padrão: `Relatorio_Forense_Final.html`. |
| `-AtivarQuarentena` | `Switch` | Não | Quando informado, isola e compacta os arquivos detectados. |
| `-SenhaZip` | `String` | Não | Senha utilizada para criptografar o pacote de evidências 7-Zip. |

---

## 🛡️ Segurança e Conformidade

Este projeto foi construído alinhado às boas práticas de **Governança da Segurança da Informação (ISO/IEC 27001)** e normas de **Privacidade de Dados (LGPD / GDPR)**, focado no suporte a investigações de segurança interna e pronta resposta a incidentes de vazamento de dados ou infecções por malware.

---
*Desenvolvido por **Marcelo Soares** | Especialista em Segurança da Informação e Computação Forense.*
