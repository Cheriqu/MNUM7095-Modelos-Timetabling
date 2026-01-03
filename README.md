# Modelos de Timetabling e Otimização (MNUM7095) - PPGMNE UFPR

Este repositório contém as implementações desenvolvidas para a disciplina de **Modelos de Timetabling** do Mestrado em Métodos Numéricos em Engenharia. O foco é a modelagem matemática e resolução exata de problemas complexos de agendamento e alocação de recursos.



## 🧩 Problemas Abordados

### 1. High School Timetabling (HSTT) - Instância Dorneles
Resolução do problema de grade horária escolar, garantindo que professores, turmas e salas sejam alocados sem conflitos, respeitando restrições pedagógicas.

* **Abordagem:** Programação Linear Inteira (PLI).
* **Restrições Hard/Soft:**
    * Conflitos de horário (Professor não pode estar em dois lugares).
    * **Aulas Duplas:** Garantia de blocos contíguos de aula para a mesma disciplina.
    * Janelas e preferências de disponibilidade.
* **Outputs:** A solução gera a grade ótima exportada para planilhas e arquivos de alocação.

### 2. Escala de Motoristas (Driver Scheduling)
Otimização da alocação de motoristas, possivelmente minimizando custos ou maximizando a cobertura de turnos, sujeito a leis trabalhistas e disponibilidade.

## 🛠️ Tecnologias e Solvers

* **Linguagem:** Python (Modelagem via bibliotecas como `pulp` ou `mip`).
* **Formatos de Otimização:**
    * `.lp`: Arquivo padrão de Programação Linear (legível por solvers como CPLEX, Gurobi, CBC).
    * `.mst`: Arquivo de solução (MIP Start) para acelerar a convergência do solver.
    * `.json`: Estrutura de dados para entrada/saída de instâncias.

## 📂 Estrutura de Arquivos

* `hstt_model_v2.ipynb`: Notebook principal com a modelagem do problema de grade escolar.
* `modelo_dorneles.lp`: O modelo matemático gerado, contendo todas as variáveis de decisão e restrições.
* `solucao_grade_horaria.xlsx - Grade Horária.csv`: A grade final gerada pelo otimizador.
* `minDoubleLessons.dat`: Dados de entrada definindo as exigências de aulas duplas.

---
**Autor:** Luiz Henrique Barretta Francisco
*Graduado em Estatística / Mestrando em Métodos Numéricos em Engenharia - UFPR*
