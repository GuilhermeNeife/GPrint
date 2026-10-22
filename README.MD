# 📖 GPrint — Sistema Integrado de Imposição e Envio para POD

Sistema desktop desenvolvido em Python para automatizar processos de pré-impressão, imposição de PDFs e envio para impressoras de rede, com registro de operações em banco de dados SQLite.

O GPrint centraliza etapas do fluxo de produção de livros em Print on Demand (POD), reduzindo tarefas manuais e facilitando o controle de impressões, reprocessamentos e histórico operacional.

## 🎯 Visão geral

O sistema integra diferentes etapas do processo produtivo:

* **Imposição de PDFs:** montagem de capas e miolos conforme regras de acabamento.
* **Preparação de arquivos:** aplicação de sangria, marcas de corte, identificação e marca d'água.
* **Envio para impressoras:** comunicação via LPR com equipamentos de rede.
* **Reimpressão e retrabalho:** seleção de arquivos com defeitos e registro de perdas.
* **Histórico de produção:** armazenamento e consulta de operações utilizando SQLite.
* **Interface desktop:** organização das funcionalidades em abas independentes.

## 🧩 Arquitetura do sistema

```
                         GPrint.py
                             |
        ┌────────────────────┼────────────────────┐
        │                    │                    │
   ┌────▼────┐          ┌────▼────┐          ┌────▼────┐
   │  CAPA   │          │  MIOLO  │          │ HISTÓRICO│
   └────┬────┘          └────┬────┘          └────┬────┘
        │                    │                    │
        └────────────────────┼────────────────────┘
                             │
              ┌──────────────┼──────────────┐
              │              │              │
         ┌────▼────┐    ┌────▼────┐    ┌────▼────┐
         │EQUIPAM. │    │ REFAZER │    │CONFIG.  │
         └────┬────┘    └────┬────┘    └─────────┘
              │              │
              └──────┬───────┘
                     │
          ┌──────────┼───────────┐
          │          │           │
     ┌────▼────┐ ┌───▼────┐ ┌────▼────┐
     │ PDFs    │ │ LPR    │ │ SQLite  │
     │ PyMuPDF │ │ Rede   │ │ Logs    │
     └─────────┘ └────────┘ └─────────┘
```

## 📑 Funcionalidades

### 1. 🟠 CAPA — Processamento de capas

Processa arquivos TXT de lotes e gera PDFs agrupados por pedido e características de produção.

**Principais recursos:**

* Leitura de arquivos TXT com campos separados por `<span>|</span>`.
* Aplicação de marca d'água com identificação do arquivo e do pedido.
* Diferenciação visual entre laminação fosca e brilho.
* Ajuste e posicionamento de PDFs em formatos de produção predefinidos.
* Agrupamento e mesclagem de capas por identificador e tipo de laminação.
* Envio para a impressora RICOH via LPR.
* Opção para enviar somente uma cópia durante reprocessamentos.
* Registro das montagens e dos envios no banco de dados.

**Formatos contemplados:** 28×56, 32×60, 32×66 e 32×48, além dos formatos adicionais previstos na configuração do sistema.

### 2. 📄 MIOLO — Imposição de miolos

Responsável pela preparação dos arquivos internos dos livros, aplicando regras de imposição conforme o código de acabamento.


| Código | Modo de imposição           | Descrição                                       |
| ------- | ----------------------------- | ------------------------------------------------- |
| 1       | Dois na folha / corta-empilha | Seleção conforme quantidade                     |
| 2       | Grampo                        | Imposição com frente e verso alternados         |
| 3       | Um na folha                   | Uma página por folha                             |
| 4       | Corta-empilha rotacionado     | Montagem de caderno rotacionado                   |
| 5       | Um na folha ATEXCO            | Orientação forçada com rotações específicas |
| 6       | Dois na folha rotacionado     | Duas páginas por folha, com rotação da segunda |

**Recursos de preparação:**

* Sangria externa espelhada, com padrão de 3 mm.
* Suporte a rotações de 0°, 90°, 180° e 270°.
* Marcas de corte e marca central de dobra.
* Definição de TrimBox e BleedBox por página.
* Identificação textual com número da página, total de páginas e nome do arquivo.
* Ajuste do posicionamento da identificação.
* Compressão opcional por Ghostscript.
* Salvamento local ou envio automático para a impressora selecionada.
* Registro independente de montagem e envio.

**Equipamentos de destino:** ATEXCO, BALTORO, OCÉ 1, OCÉ 3 e RICOH.

### 3. 📊 HISTÓRICO — Consulta de produção

Interface para consultar os registros de impressão e acompanhar o histórico operacional.

**Filtros disponíveis:**

* Intervalo de datas.
* Lote, pilha e arquivo.
* Máquina e usuário.
* Status da operação.
* Modo de imposição.

**Recursos adicionais:**

* Exibição de até 5.000 registros.
* Ordenação por colunas.
* Soma das quantidades registradas.
* Consulta ao banco compartilhado, com alternativa local em caso de indisponibilidade da rede.

### 4. 🖨️ EQUIPAMENTOS — Envio de PDFs

Permite enviar arquivos PDF já preparados diretamente para os equipamentos de impressão, sem executar novamente o processo de imposição.

**Funcionalidades:**

* Seleção manual e arrastar e soltar arquivos.
* Remoção individual ou limpeza da lista.
* Seleção de máquina, formato, papel e gramatura.
* Identificação opcional do trabalho.
* Montagem automática do nome da fila LPR.
* Processamento em thread separada.
* Barra de progresso e opção de cancelamento.

**Equipamentos configurados:**


| Máquina | Endereço IP                  | Grupo |
| -------- | ----------------------------- | ----- |
| OCÉ 1   | `<span>192.168.200.80</span>` | PB    |
| OCÉ 3   | `<span>192.168.200.85</span>` | PB    |
| RICOH    | `<span>192.168.200.84</span>` | CL    |
| BALTORO  | `<span>192.168.200.73</span>` | CL    |

Os formatos e as filas disponíveis dependem da configuração de cada equipamento.

### 5. 🔄 REFAZER — Gestão de reimpressões

Módulo destinado ao tratamento de arquivos com defeito e ao controle de retrabalho.

**Fluxo operacional:**

1. Seleção do cliente e do lote.
2. Localização da planilha de produção `<span>.xlsx</span>`.
3. Leitura das informações de capas, miolos, caminhos, formatos e acabamentos.
4. Seleção dos arquivos que precisam ser refeitos.
5. Definição da quantidade, setor de perda, máquina e operador.
6. Preparação dos arquivos na estrutura `<span>REFAZER/CAPA</span>` ou `<span>REFAZER/MIOLO</span>`.
7. Registro do retrabalho na planilha de perdas.

**Tratamentos aplicados:**

* Marca d'água vermelha nas capas refeitas.
* Conversão de fontes em contornos via Ghostscript, quando necessário.
* Imposição de miolos conforme o acabamento.
* Compressão de arquivos para casos de pré-impressão.
* Registro de quantidade, setor, máquina, operador, data e tipo de arquivo.

A planilha de perdas utilizada é `<span>Z_RELATÓRIOS/RELATÓRIOS DE PERDA.xlsx</span>`, com registros nas abas `<span>UICLAP</span>` ou `<span>PODS</span>`.

**Setores de perda contemplados:**

Pré-impressão, Problema de fonte, Impressão, Refile inicial miolo, Refile inicial capa, Separação, Encadernação, Laminação, Refile orelha, Refile final, Shrink e Expedição.

### 6. ⚙️ CONFIGURAÇÕES

Permite personalizar parâmetros utilizados no processamento dos arquivos.

**Opções disponíveis:**

* Sangria padrão dos miolos.
* Cores da marca d'água para laminação fosca e brilho.
* Fonte e tamanho da marca d'água.
* Deslocamentos horizontal e vertical.
* Deslocamento específico para o formato 28×56.
* Restauração dos parâmetros padrão.
* Seleção de temas claros e escuros.

As configurações são persistidas em:

```
~/uiclap_config.json
```

Temas disponíveis por meio do `<span>ttkbootstrap</span>` incluem Cosmo, Flatly, Journal, Litera, Lumen, Minty, Pulse, Sandstone, United, Yeti, Morph, Darkly, Superhero, Solar, Cyborg e Vapor.

## 🗄️ Banco de dados e rastreabilidade

O GPrint utiliza SQLite para registrar as operações de produção em uma tabela denominada `<span>logs</span>`.

**Principais informações registradas:**


| Campo                      | Finalidade                        |
| -------------------------- | --------------------------------- |
| `<span>data_hora</span>`   | Data e hora da operação         |
| `<span>lote</span>`        | Caminho da pasta do lote          |
| `<span>pilha</span>`       | Identificação do TXT processado |
| `<span>arquivo</span>`     | Nome do arquivo                   |
| `<span>maquina</span>`     | Equipamento de destino            |
| `<span>ip</span>`          | Endereço IP do equipamento       |
| `<span>formato</span>`     | Formato de impressão             |
| `<span>modo</span>`        | Método de imposição            |
| `<span>quantidade</span>`  | Quantidade processada             |
| `<span>status</span>`      | Resultado da operação           |
| `<span>erro</span>`        | Detalhes de falhas                |
| `<span>copia</span>`       | Identificação da cópia         |
| `<span>usuario</span>`     | Usuário responsável             |
| `<span>montagem_ok</span>` | Indicador de sucesso da montagem  |
| `<span>envio_ok</span>`    | Indicador de sucesso do envio     |

**Mecanismos de proteção:**

* Arquivos de bloqueio para controle de acesso concorrente.
* Uso de `<span>O_EXCL</span>` na criação dos arquivos de trava.
* Modo WAL (`<span>Write-Ahead Logging</span>`) do SQLite.
* Banco alternativo local quando o compartilhamento de rede está indisponível.

O banco principal é utilizado pelas abas de produção, retrabalho e histórico, centralizando a rastreabilidade das operações.

## 🔁 Fluxo de produção

```
          TXT DO LOTE
               |
               ▼
       Leitura e validação
               |
               ▼
       Identificação do modo
               |
               ▼
        Preparação do PDF
       ┌───────┼────────┐
       │       │        │
    Sangria  Marcas  Identificação
       │       │        │
       └───────┼────────┘
               |
               ▼
        Imposição final
               |
               ▼
       ┌───────┴────────┐
       │                │
       ▼                ▼
  Salvar PDF        Enviar via LPR
       │                │
       └───────┬────────┘
               |
               ▼
        Registro no SQLite
```

## 🛠️ Tecnologias e dependências


| Tecnologia                    | Aplicação                                     |
| ----------------------------- | ----------------------------------------------- |
| Python                        | Linguagem principal                             |
| Tkinter                       | Interface desktop                               |
| ttkbootstrap                  | Temas e componentes visuais                     |
| PyMuPDF (`<span>fitz</span>`) | Manipulação de PDFs                           |
| PyPDF2                        | Mesclagem de documentos                         |
| pikepdf                       | Operações de baixo nível em PDFs             |
| ReportLab                     | Geração de elementos e páginas PDF           |
| Pillow                        | Processamento de imagens e sangria              |
| openpyxl                      | Leitura e escrita de planilhas Excel            |
| tkinterdnd2                   | Funcionalidade de arrastar e soltar             |
| SQLite                        | Persistência e histórico                      |
| Ghostscript                   | Compressão e conversão de fontes em contornos |
| LPR                           | Envio de trabalhos para impressoras de rede     |

### Requisitos externos

* Windows com suporte ao serviço e cliente LPR.
* Ghostscript instalado e acessível no caminho esperado pelo sistema.
* Acesso à rede de produção e aos equipamentos configurados.
* Permissões de leitura e escrita nas pastas e planilhas utilizadas.
* Filas de impressão configuradas com os nomes esperados pelo sistema.

## ⚠️ Observações técnicas

* **Infraestrutura de rede:** existem caminhos fixos para compartilhamentos e recursos internos de produção.
* **Impressão:** o envio depende da disponibilidade do serviço LPR e da correspondência entre os nomes de filas configurados e os utilizados pelo programa.
* **Ghostscript:** determinadas operações dependem da instalação e localização corretas do executável.
* **Concorrência:** operações pesadas são executadas em threads separadas para preservar a responsividade da interface.
* **Banco de dados:** o compartilhamento SQLite exige atenção à disponibilidade da rede, às permissões e ao controle de concorrência.

## 🧭 Casos de uso


| Necessidade                      | Módulo         |
| -------------------------------- | --------------- |
| Processar um lote de capas       | CAPA            |
| Impor miolos de livros           | MIOLO           |
| Reenviar PDFs prontos            | EQUIPAMENTOS    |
| Refazer arquivos com defeito     | REFAZER         |
| Consultar impressões anteriores | HISTÓRICO      |
| Ajustar sangria e marca d'água  | CONFIGURAÇÕES |

## 👨‍💻 Autor

**Guilherme Neife**

**Projeto:** GPrint — v10.0

Sistema desenvolvido para automatizar e organizar o fluxo de produção de livros em Print on Demand (POD), com foco em imposição de PDFs, integração com impressoras de rede e rastreabilidade operacional.

---

*Documentação baseada nas funcionalidades descritas do projeto. Caminhos, endereços IP, dependências e comportamentos operacionais devem ser conferidos na versão efetivamente distribuída do código.*
