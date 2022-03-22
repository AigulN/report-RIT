# report-for-RIT.py
import click
import numpy as np
import pandas as pd
#from datetime import date
import holidays
import matplotlib.pyplot as plt


income = 24000
xlsx_TimeTask = pd.read_excel("App_A.xlsx", sheet_name="App_A")
xlsx_Score = pd.read_excel("App_B.xlsx", sheet_name="App_B")
xlsx_Rate = pd.read_excel("App_C.xlsx", sheet_name="App_C")


@click.command()
def main():
    tasks = xlsx_TimeTask["Задача"].unique()
    performers = xlsx_TimeTask["Исполнитель"].unique()
    performers.sort()
    days = xlsx_TimeTask["Дата"].unique()
    days.sort()
    days = pd.to_datetime(days)
    start_day = days[0]
    end_day = days[days.size - 1]
    xlsx_TimeTask["Год"] = pd.DatetimeIndex(xlsx_TimeTask['Дата']).year
    years = xlsx_TimeTask['Год'].unique()
    list_days = pd.date_range(start_day, end_day)
    ru_holidays = holidays.country_holidays("RU", years=years)
    list_workdays = pd.to_datetime([x for x in list_days if x not in ru_holidays])
    expenses_perfs = []
    mean_hours_perfs = []
    absence_days = []
    outliers = []

    print("Общие трудозатраты на проект: " + str(xlsx_TimeTask["Часы"].sum()) + " часов")
    hours_on_task = []
    for x in tasks:
        hours_on_task.append(xlsx_TimeTask["Часы"][xlsx_TimeTask["Задача"] == x].sum())
    print("Среднее время, затраченное на решение задач: " + str(round(np.mean(hours_on_task), 1)) + " часов")
    print("Медианное время, затраченное на решение задач: " + str(round(np.median(hours_on_task), 1)) + " часов")
    #hours_on_performers = []

    df_outliers = pd.DataFrame({"Исполнитель": performers})
    for x in performers:
        sum_outliers = 0
        df_x = xlsx_TimeTask[["Дата", "Часы", "Задача"]][xlsx_TimeTask["Исполнитель"] == x]
        expenses_perfs.append(df_x["Часы"].sum() * int(xlsx_Rate["Ставка"][xlsx_Rate["Исполнитель"] == x]))
        #dop_mas = xlsx_TimeTask["Часы"][xlsx_TimeTask["Исполнитель"] == x]
        #expenses_perfs.append(dop_mas.sum() * int(xlsx_Rate["Ставка"][xlsx_Rate["Исполнитель"] == x]))
        dates_x = pd.to_datetime(df_x["Дата"].unique())
        tasks_x = df_x["Задача"].unique()
        hours_x_on_task = []
        for y in dates_x:
            hours_x_on_task.append(df_x["Часы"][df_x["Дата"] == y].sum())
        mean_hours_perfs.append(np.mean(hours_x_on_task))
        absence_days_x = [str(d.date()) for d in list_workdays if d not in df_x["Дата"].unique()]
        absence_days.append(absence_days_x)
        #hours_on_performers.append(dop_mas.sum())
        #print("Среднее время, затраченное на решение задач исполнителем " + x + ": "+ str(round(dop_mas.mean(), 1)) + " часов")
        sumhours_task_x = []
        for y in tasks_x:
            sumhours_task_x.append(df_x["Часы"][df_x["Задача"] == y].sum())
        df_task_x = pd.DataFrame({"Задача": tasks_x, "Время": sumhours_task_x})
        print("Среднее время, затраченное на решение задач исполнителем " + x + ": " + str(round(df_task_x["Время"].mean(), 1)) + " часов")
        df_merge_x = pd.merge(df_task_x, xlsx_Score)
        df_merge_x["Время-Оценка"] = df_merge_x["Время"] - df_merge_x["Оценка"]
        df_merge_x["Вылет"] = df_merge_x["Время-Оценка"] / df_merge_x["Оценка"]
        #outliers.append(int(df_merge_x["Вылет"][df_merge_x["Вылет"] > 0].mean()*100))
        outliers.append(int(df_merge_x["Вылет"][df_merge_x["Вылет"] > 0].sum()/df_merge_x['Вылет'][df_merge_x["Вылет"] >= 0].shape[0] * 100))

    df_outliers["Вылет"] = outliers

    Profit = (income - np.sum(expenses_perfs)) * 100 / income
    print("Рентабельность проекта: " + str(round(Profit, 1)))

    df_hours_on_performers = pd.DataFrame({'Исполнитель': performers, 'Время': mean_hours_perfs})
    for x in performers:
        print("Среднее количество часов, отрабатываемое за день сотрудником " + x + ": " + str(round(df_hours_on_performers["Время"][df_hours_on_performers["Исполнитель"] == x].sum(), 1)))

    for i, x in enumerate(performers):
        if len(absence_days[i]) != 0:
            print(f"Дни отсутствия сотрудника {x}:") #+ str(absence_days[i]))
            print(*absence_days[i], sep=", ")
        else:
            print(f"Сотрудник {x} не имеет дней отсутствия")

    for x in performers:
        print("Средний \"вылет\" из оценки специалиста " + x + ": " + str(df_outliers["Вылет"][df_outliers["Исполнитель"] == x].sum()) + " %")

    totalHours_onTask = []
    for x in xlsx_Score["Задача"]:
        totalHours_onTask.append(xlsx_TimeTask["Часы"][xlsx_TimeTask["Задача"] == x].sum())
    xlsx_Score["Факт"] = totalHours_onTask
    fig, ax = plt.subplots()
    fig.set_figheight(8)
    fig.set_figwidth(25)
    index = np.arange(len(xlsx_Score["Задача"]))
    bar_width = 0.4
    opacity = 0.5
    st_score = plt.bar(index, xlsx_Score["Оценка"], bar_width, alpha=opacity, color='y', label='Оценка')
    st_fact = plt.bar(index + bar_width, xlsx_Score["Факт"], bar_width, alpha=opacity, color='b', label='Факт')
    plt.xlabel("Задача LOC-")
    plt.ylabel("Часы")
    plt.title("Оценки трудозатрат и их фактические значения")
    plt.legend()
    plt.xticks(index + bar_width / 2, index + 1)
    fig.savefig("Score_fact.png")


if __name__ == "__main__":
    main()