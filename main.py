import numpy as np
import pandas as pd
import holidays
import matplotlib.pyplot as plt
import warnings


income = 24000
fileA = 'App_A.xlsx'
sheetA = 'App_A'
fileB = 'App_B.xlsx'
sheetB = 'App_B'
fileC = 'App_C.xlsx'
sheetC = 'App_C'
warnings.filterwarnings("ignore", category=np.VisibleDeprecationWarning)


def read_file(fileNameA, sheetNameA, fileNameB, sheetNameB, fileNameC, sheetNameC):
    TimeTask = pd.read_excel(fileNameA, sheet_name=sheetNameA)
    Score = pd.read_excel(fileNameB, sheet_name=sheetNameB)
    Rate = pd.read_excel(fileNameC, sheet_name=sheetNameC)
    return TimeTask, Score, Rate


def calculate_marks(xlsx_TimeTask, xlsx_Score, xlsx_Rate):
    df_project = pd.DataFrame()
    df_performers = pd.DataFrame(
        {"Исполнитель": [], "Ср.время на задачу": [], "Часов в день": [], "Список пропущ-х дней": [],
         "Ср.вылет из оценки": []})
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
    ru_holidays = holidays.country_holidays("RU", years=years)
    list_days = pd.date_range(start_day, end_day)
    list_workdays = pd.to_datetime([x for x in list_days if x not in ru_holidays])

    expenses_perfs = []
    hours_on_task = []

    for x in tasks:
        hours_on_task.append(xlsx_TimeTask["Часы"][xlsx_TimeTask["Задача"] == x].sum())
    df_project["Время, затраченное на решение каждой задачи"] = hours_on_task

    for x in performers:
        df_x = xlsx_TimeTask[["Дата", "Часы", "Задача"]][xlsx_TimeTask["Исполнитель"] == x]
        expenses_perfs.append(df_x["Часы"].sum() * int(xlsx_Rate["Ставка"][xlsx_Rate["Исполнитель"] == x]))
        dates_x = pd.to_datetime(df_x["Дата"].unique())
        tasks_x = df_x["Задача"].unique()
        hours_x_on_task = []
        for y in dates_x:
            hours_x_on_task.append(df_x["Часы"][df_x["Дата"] == y].sum())
        absence_days_x = [str(d.date()) for d in list_workdays if d not in df_x["Дата"].unique()]
        sum_hours_task_x = []
        for y in tasks_x:
            sum_hours_task_x.append(df_x["Часы"][df_x["Задача"] == y].sum())
        df_task_x = pd.DataFrame({"Задача": tasks_x, "Время": sum_hours_task_x})
        df_merge_x = pd.merge(df_task_x, xlsx_Score)
        df_merge_x["Вылет"] = (df_merge_x["Время"] - df_merge_x["Оценка"]) / df_merge_x["Оценка"]
        outlier = int(df_merge_x["Вылет"][df_merge_x["Вылет"] > 0].sum() / len(df_merge_x.index) * 100)
        df_performers.loc[len(df_performers.index)] = [x, round(df_task_x["Время"].mean(), 1), np.mean(hours_x_on_task),
                                                       absence_days_x, outlier]
    df_project["Рентабельность"] = (income - np.sum(expenses_perfs)) * 100 / income
    return df_project, df_performers


def getbar(BarData1, BarData2, labelbar1, labelbar2, labelx, labely, bartitle, filename):
    fig, ax = plt.subplots()
    fig.set_figheight(8)
    fig.set_figwidth(25)
    index = np.arange(len(BarData1))
    bar_width = 0.4
    opacity = 0.5
    st_score = plt.bar(index, BarData1, bar_width, alpha=opacity, color='y', label=labelbar1)
    st_fact = plt.bar(index + bar_width, BarData2, bar_width, alpha=opacity, color='b', label=labelbar2)
    plt.xlabel(labelx)
    plt.ylabel(labely)
    plt.title(bartitle)
    plt.legend()
    plt.xticks(index + bar_width / 2, index + 1)
    fig.savefig(filename)


def main(fileNameA, sheetNameA, fileNameB, sheetNameB, fileNameC, sheetNameC):

    xlsx_TimeTask, xlsx_Score, xlsx_Rate = read_file(fileNameA, sheetNameA, fileNameB, sheetNameB, fileNameC,
                                                     sheetNameC)
    mydfProject, mydfPerformers = calculate_marks(xlsx_TimeTask, xlsx_Score, xlsx_Rate)

    print("Общие трудозатраты на проект: " + str(xlsx_TimeTask["Часы"].sum()) + " часов")
    print("Среднее время, затраченное на решение задач: " + str(
        round(mydfProject["Время, затраченное на решение каждой задачи"].mean(), 1)) + " часов")
    print("Медианное время, затраченное на решение задач: " + str(
        round(mydfProject["Время, затраченное на решение каждой задачи"].median(), 1)) + " часов")
    performers = mydfPerformers["Исполнитель"]
    for x in performers:
        print("Среднее время, затраченное на решение задач исполнителем " + x + ": " + str(
            mydfPerformers["Ср.время на задачу"][mydfPerformers["Исполнитель"] == x].sum()) + " часов")
    print("Рентабельность проекта: " + str(round(mydfProject["Рентабельность"].values[0], 1)))
    for x in performers:
        print("Среднее количество часов, отрабатываемое за день сотрудником " + x + ": " + str(
            round(mydfPerformers["Часов в день"][mydfPerformers["Исполнитель"] == x].sum(), 1)))

    for x in performers:
        if len(*mydfPerformers["Список пропущ-х дней"][mydfPerformers["Исполнитель"] == x]) != 0:
            print(f"Дни отсутствия сотрудника {x}:")
            print(*mydfPerformers["Список пропущ-х дней"][mydfPerformers["Исполнитель"] == x])
        else:
            print(f"Сотрудник {x} не имеет дней отсутствия")

    for x in performers:
        print("Средний \"вылет\" из оценки специалиста " + x + ": " + str(
            mydfPerformers["Ср.вылет из оценки"][mydfPerformers["Исполнитель"] == x].sum()) + " %")

    totalhours_onTask = []
    for x in xlsx_Score["Задача"]:
        totalhours_onTask.append(xlsx_TimeTask["Часы"][xlsx_TimeTask["Задача"] == x].sum())
    xlsx_Score["Факт"] = totalhours_onTask

    getbar(xlsx_Score["Оценка"], xlsx_Score["Факт"], "Оценка", "Факт", "Задача LOC-", "Часы",
           "Оценки трудозатрат и их фактические значения", "Score_fact.png")


if __name__ == '__main__':
    main(fileA, sheetA, fileB, sheetB, fileC, sheetC)
