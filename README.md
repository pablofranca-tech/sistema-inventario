# 🏭 Sistema de Inventário para Multinacional

![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-production-brightgreen)

Sistema completo para gestão de inventário industrial com controle de paletes e progressivas, desenvolvido para as unidades de Campinas e Mafra.

## ✨ Funcionalidades Principais

### 📷 Módulo de Leitura
- **Registro inteligente** de etiquetas de paletes e progressivas
- Validação em tempo real do formato das etiquetas
- Verificação automática de duplicatas e etiquetas não vinculadas
- Busca e comparação com códigos existentes nas planilhas das pastas PRODUÇÃO e RETRABALHO
- Alerta de inconsistências entre leitura e base de dados
- Visualização interativa dos códigos a serem lidos

### 📁 Módulo de Consulta
- Pesquisa avançada com 5 filtros simultâneos
- Histórico completo de registros
- Controle de usuários e plantas
- Destaque visual para:
  - ⚠️ Etiquetas repetidas
  - ❌ Itens não apontados
  - ✅ Registros válidos

### 📊 Módulo de Análise
- Dashboard com métricas em tempo real
- Gráficos profissionais:
  - Distribuição por tipo de etiqueta
  - Comparativo únicas vs repetidas
  - Evolução diária de registros
  - Ranking por operador
- Filtros temporais flexíveis

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Finalidade | Versão |
|------------|------------|--------|
| Python | Linguagem principal | 3.8+ |
| CustomTkinter | Interface moderna | 5.2 |
| Pandas | Processamento de dados | 1.3+ |
| Matplotlib | Visualização de dados | 3.5+ |
| OpenPyXL | Integração com Excel | 3.0+ |

