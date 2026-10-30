# GPrint — Sistema de Automação de Impressão e Processamento de PDFs

**Python | Automação de Processos | Processamento de PDFs | SQLite | Integração com Impressoras Industriais**

## Sobre o projeto

O **GPrint** é uma aplicação desktop desenvolvida em Python para automatizar o processamento, a montagem e o envio de arquivos PDF para equipamentos de impressão digital em um ambiente de produção gráfica.

O sistema foi criado para otimizar atividades que anteriormente exigiam diversas intervenções manuais, centralizando em uma única interface a preparação dos documentos, a configuração dos trabalhos, o envio às impressoras e o acompanhamento das operações.

Atualmente, é uma **solução funcional utilizada em ambiente produtivo real**, operando simultaneamente em quatro estações de trabalho, com possibilidade de expansão para outros computadores.

O projeto foi desenvolvido e aprimorado continuamente, considerando necessidades operacionais, desempenho, estabilidade e facilidade de utilização pelos profissionais da produção.

## Principais funcionalidades

### Processamento e montagem automatizada de PDFs

O GPrint realiza a leitura e o processamento de documentos PDF, aplicando diferentes modalidades de montagem conforme as características de cada trabalho.

Entre as funcionalidades estão:

* Montagem automatizada de arquivos de capa e miolo.
* Imposição de páginas para diferentes processos de impressão.
* Montagem nos formatos corta e empilha, grampo e múltiplas páginas por folha.
* Rotação e posicionamento automático das páginas.
* Aplicação de sangria, margens e marcas de corte.
* Tratamento de diferentes formatos e dimensões.
* Mesclagem de arquivos PDF para otimização dos trabalhos.
* Organização e gerenciamento de arquivos temporários e finais.

### Integração e envio para impressoras

O sistema permite o envio automatizado dos arquivos processados diretamente para equipamentos de impressão digital por meio do protocolo **LPR (Line Printer Remote)**.

Entre os recursos implementados estão:

* Comunicação com diferentes impressoras de rede.
* Seleção de equipamentos e filas de impressão.
* Envio automatizado de trabalhos após o processamento.
* Controle de quantidades e cópias.
* Identificação dos trabalhos enviados.
* Tratamento de erros de comunicação e controle de tempo de resposta.
* Acompanhamento do resultado das transmissões.

Essa integração reduz intervenções manuais e permite que o operador execute diferentes etapas diretamente pela aplicação.

### Processamento de lotes

O GPrint permite importar arquivos TXT estruturados que contêm as informações necessárias para identificar e organizar os trabalhos de produção.

O sistema interpreta esses dados, identifica os documentos correspondentes e executa os processos de maneira sequencial e automatizada.

Também oferece suporte à seleção de múltiplos arquivos de entrada, com controle configurável da quantidade de lotes processados em uma execução.

### Histórico e rastreabilidade

Foi implementado um sistema de registros utilizando **SQLite**, permitindo acompanhar as operações realizadas.

Os dados registrados incluem:

* Data e horário das operações.
* Identificação de lotes e arquivos.
* Usuário responsável.
* Impressora e fila utilizada.
* Modalidade de processamento.
* Quantidade de trabalhos e cópias.
* Situação da montagem e do envio.
* Ocorrências e falhas identificadas.

O sistema também conta com mecanismos de controle de gravações simultâneas, armazenamento local de registros pendentes e sincronização com o histórico central.

A interface de consulta disponibiliza filtros e indicadores que ajudam a acompanhar a produtividade e identificar atividades realizadas por equipamento, operador ou lote.

### Interface e monitoramento

A aplicação utiliza uma interface gráfica desenvolvida com **Tkinter e ttkbootstrap**, organizada em abas para facilitar a operação.

Entre os recursos estão:

* Configurações personalizáveis de processamento.
* Visualização dos arquivos e trabalhos selecionados.
* Barras de progresso.
* Logs de execução.
* Monitoramento da memória RAM do computador.
* Cronômetro de processamento.
* Indicadores visuais de status.
* Consulta ao histórico de produção.

## Arquitetura e tecnologias


| Tecnologia             | Aplicação                                            |
| ---------------------- | ------------------------------------------------------ |
| Python                 | Desenvolvimento principal e automação                |
| Tkinter / ttkbootstrap | Interface gráfica desktop                             |
| PyMuPDF / pikepdf      | Manipulação e processamento de PDFs                  |
| ReportLab              | Geração de elementos gráficos em PDFs               |
| SQLite                 | Persistência e histórico operacional                 |
| OpenPyXL               | Manipulação de planilhas Excel                       |
| LPR                    | Comunicação com impressoras de rede                  |
| Threading              | Execução de tarefas sem bloquear a interface         |
| Subprocess             | Integração com ferramentas e comandos externos       |
| PyInstaller            | Distribuição da aplicação como executável Windows |
| Git / GitHub           | Versionamento e manutenção do código                |

## Utilização em ambiente produtivo

O GPrint é utilizado por **quatro operadores em quatro computadores diferentes**, permitindo a execução simultânea de tarefas.

A aplicação é disponibilizada por meio de um executável em rede, enquanto o processamento dos documentos utiliza os recursos de cada estação de trabalho.

Essa arquitetura permite executar operações em diferentes equipamentos, mantendo os registros em um histórico compartilhado.

Durante seu desenvolvimento, foram considerados aspectos como:

* Processamento de arquivos PDF de diferentes tamanhos e complexidades.
* Integração entre aplicações e equipamentos físicos.
* Execução de operações simultâneas.
* Tratamento de falhas e indisponibilidade de recursos de rede.
* Organização e limpeza de arquivos temporários.
* Rastreabilidade das operações.
* Estabilidade e continuidade das rotinas produtivas.

## Resultados e benefícios

O desenvolvimento do GPrint contribuiu para:

* **Automatização de tarefas repetitivas**, reduzindo a necessidade de operações manuais.
* **Padronização dos processos de impressão**, centralizando diferentes modalidades de montagem.
* **Agilidade operacional**, automatizando a preparação e o envio dos arquivos.
* **Maior rastreabilidade**, com registros de processamento e transmissão dos trabalhos.
* **Redução da complexidade operacional**, permitindo executar diferentes tarefas em uma interface centralizada.
* **Melhoria contínua**, com novas funcionalidades desenvolvidas a partir de necessidades reais dos usuários.

## Minha atuação no desenvolvimento

Fui responsável pelo desenvolvimento e evolução do GPrint, atuando desde a identificação das necessidades operacionais até a implementação das funcionalidades e manutenção da aplicação.

Entre as atividades realizadas estão:

* Análise e levantamento de requisitos.
* Desenvolvimento de aplicações desktop em Python.
* Implementação de algoritmos de processamento e montagem de documentos.
* Integração com equipamentos de impressão digital.
* Implementação de persistência de dados com SQLite.
* Desenvolvimento de interfaces gráficas.
* Tratamento de erros e melhorias de desempenho.
* Testes, correções e manutenção.
* Suporte aos usuários e evolução contínua do sistema.

## Conclusão

O GPrint representa a aplicação prática dos meus conhecimentos em **Python, desenvolvimento de software, automação e análise de processos operacionais**.

Mais do que automatizar uma atividade específica, o projeto demonstra minha capacidade de compreender necessidades de um ambiente produtivo, desenvolver soluções técnicas e manter uma aplicação utilizada diariamente por diferentes profissionais.

É um projeto que continua evoluindo, com foco em eficiência, confiabilidade, organização e melhoria contínua dos processos.

---

**Desenvolvido por Guilherme Neife**

*Python Developer | Automação de Processos | Soluções Tecnológicas*
