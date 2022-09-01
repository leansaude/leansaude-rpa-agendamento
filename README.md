# RPA de agendamento de visitas Lean Stay
Identifica visitas médicas pendentes de agendamento e realiza os agendamentos no Amplimed.

## Observações
1. Ao clonar o repositório, não esqueça de criar um arquivo .env baseado no exemplo contido em .env.exemplo. 
Ajuste as configurações como desejar.

2. Este repositório faz uso de virtual envs do python. Para usá-lo, você precisa ter a seguinte dependência instalada globalmente em seu computador:
```
pip install virtualenv
```

3. As dependências específicas deste projeto estão descritas em requirements.txt. 
Instale-as rodando:
```
cd <caminho-do-seu-repo>
env\Scripts\pip install -r requirements.txt
```

4. Execute o script usando o python contido no virtual env, e não o python global.
```
cd <caminho-do-seu-repo>
env\Scripts\python agendamento.py
```