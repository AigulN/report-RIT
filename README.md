# Установка
```
git clone https://github.com/AigulN/report-RIT.git
```
```
cd report-RIT
```
Создать и активировать виртуальную среду.
```
virtualenv venv
```
Linux или Mac:

```
source venv/bin/activate
```

Windows (выполнять в PowerShell):

```
Set-ExecutionPolicy Unrestricted -Scope Process
```
```
.\venv\Scripts\activate
```

Установить необходимые зависимости и модули:
```
pip install -r requirements.txt
```
Запустить код:
```
python main.py
```