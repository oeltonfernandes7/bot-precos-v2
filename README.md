# Bot de Preços v2.0

Sistema completo de pesquisa de preços em farmácias com interface Streamlit, logs profissionais e gestão dinâmica de fontes.

## 🚀 Como Usar

1. **Instalação**: Execute `iniciar.bat` (instala dependências automaticamente)
2. **Interface**: Abra http://localhost:8501 no navegador
3. **Pesquisa**: Importe arquivo Excel/CSV com produtos e inicie a busca

## 📁 Estrutura de Pastas

- `/logs` - Arquivos de log de cada execução
- `/historico` - CSVs de cada pesquisa separados por data
- `fontes.json` - Configuração dinâmica das fontes de pesquisa
- `app.py` - Interface Streamlit
- `bot_precos.py` - Lógica de busca de preços
- `config.py` - Configurações gerais
- `iniciar.bat` - Script de inicialização

## 🌐 Fontes Suportadas

- Farmácia Indiana (JSON VTEX)
- Drogasil (Playwright)
- Mercado Livre (Playwright)
- Pacheco (JSON VTEX)

## 📊 Funcionalidades

- **Interface Web**: 4 abas completas (Pesquisa, Histórico, Fontes, Logs)
- **Logs Profissionais**: Arquivos timestamped com níveis INFO/WARNING/ERROR
- **Gestão Dinâmica**: Fontes configuráveis via JSON
- **Histórico**: Análise de pesquisas anteriores com gráficos
- **Export**: Excel formatado com cores e formatação

## 🔧 Desenvolvimento

Para adicionar nova fonte:
1. Adicione em `fontes.json` com status "pendente"
2. Implemente método de busca em `bot_precos.py`
3. Teste e mude status para "ativa"

## 📝 Logs

Cada execução gera `logs/log_YYYYMMDD_HHMMSS.txt` com:
- Ações normais (INFO)
- Avisos de problemas não críticos (WARNING)
- Erros técnicos (ERROR)
- Resumo final com estatísticas