import matplotlib.pyplot as plt
import pandas as pd
import pyodbc as db
import numpy as np
from datetime import date
import calendar
import datetime


def convert(number):
    number = number / 1000
    number = int(number)
    number = format(number, ',')
    number = number + 'K'
    return number


connection = db.connect('DRIVER={SQL Server};'
                        'SERVER=137.116.139.217;'
                        'DATABASE=ARCHIVESKF;'
                        'UID=sa;PWD=erp@123')

def cumulative_target_sales(name):
    everday_sale_df = pd.read_sql_query("""select  right(TRANSDATE, 2) as Date, isnull(sum(QTYSHIPPED),0) as ItemSales from OESalesDetails
                                        left join PRINFOSKF
                                        on OESalesDetails.ITEM = PRINFOSKF.ITEMNO
                                        where left(TRANSDATE,6) = convert(varchar(6),getdate(), 112)
                                        and PRINFOSKF.GPMNAME like ?
                                        group by right(TRANSDATE, 2)
                                        order by right(TRANSDATE, 2)""", connection, params={name})

    day_wise_date = everday_sale_df['Date'].tolist()
    day_to_day_sale = everday_sale_df['ItemSales'].tolist()
    print('day wise date list : ', day_wise_date)
    print('day wise sales list : ', day_to_day_sale)

    from datetime import date

    today = date.today()

    current_day = today.strftime("%d")
    current_day_in_int = int(current_day)
    print("Current days in month till today: ", current_day_in_int)

    final_days_array = []
    final_sales_array = []
    final_return_array = []
    for t_va in range(0, current_day_in_int):
        final_days_array.append(day_wise_date[t_va])
        final_sales_array.append(day_to_day_sale[t_va])
    print('Reduced 0 value days from the list : ', final_days_array)
    print('Sales Taken According to the value of days : ', final_sales_array)

    EveryDay_target_df = pd.read_sql_query("""Declare @CurrentMonth NVARCHAR(MAX);
                                    Declare @DaysInMonth NVARCHAR(MAX);
                                    SET @CurrentMonth = convert(varchar(6), GETDATE(),112)
                                    SET @DaysInMonth = DAY(EOMONTH(GETDATE())) 
                                    select  isnull(sum(QTY),0)/@DaysInMonth as MonthsTargetQty 
                                    from ARCSECONDARY.dbo.RfieldForceProductTRG
                                    left join PRINFOSKF
                                    on RfieldForceProductTRG.ITEMNO = PRINFOSKF.ITEMNO
                                    where YEARMONTH=CONVERT(varchar(6), dateAdd(month,-1,getdate()), 112) and GPMNAME like ?""",
                                           connection, params={name})

    single_day_target = EveryDay_target_df['MonthsTargetQty'][0]
    # print('Single Day Target : ', single_day_target)

    y_pos = np.arange(len(final_days_array))

    n = 1
    labell_to_plot = []
    for z in y_pos:
        labell_to_plot.append(n)
        n = n + 1
    print('Final label to plot : ', labell_to_plot)

    # ----------------code for cumulitive sales------------
    import calendar
    import datetime

    now = datetime.datetime.now()
    total_days = calendar.monthrange(now.year, now.month)[1]
    # print('Total number of days in this month : ', total_days)

    final_target_day_wise = 0
    cumulative_target_that_needs_to_plot = []
    for t_value in range(0, total_days + 1):
        final_target_day_wise = single_day_target * t_value
        cumulative_target_that_needs_to_plot.append(final_target_day_wise / 1000)
        final_target_day_wise = 0
    # print('cumulative target from 0 to final day of the month : ', cumulative_target_that_needs_to_plot)

    new_array_of_cumulative_sales = [0]
    final = 0
    for val in final_sales_array:
        # print(val)
        get_in = final_sales_array.index(val)
        # print(get_in)
        if get_in == 0:
            new_array_of_cumulative_sales.append(val / 1000)
        else:
            for i in range(0, get_in + 1):
                final = (final + final_sales_array[i]) / 1000
            new_array_of_cumulative_sales.append(final)
            final = 0

    # print('Cumulative sales list from 0 to current day of month : ', new_array_of_cumulative_sales)  #
    # --------------------------sales data

    # sys.exit()

    length_of_cumulative_target = range(len(cumulative_target_that_needs_to_plot))
    length_of_cumulative_sales = range(len(new_array_of_cumulative_sales))

    list_index_for_target = len(cumulative_target_that_needs_to_plot) - 1
    # print('index size for target : ', list_index_for_target)
    # sys.exit()

    list_index_for_sale = len(new_array_of_cumulative_sales) - 1
    # print('index size for sale : ', list_index_for_sale)
    # sys.exit()

    fig, ax = plt.subplots(figsize=(5.6, 4.8))
    plt.fill_between(length_of_cumulative_sales, new_array_of_cumulative_sales, color="green", alpha=1)
    plt.plot(length_of_cumulative_target, cumulative_target_that_needs_to_plot, color="red", linewidth=3, linestyle="-")

    plt.text(list_index_for_sale + .2, new_array_of_cumulative_sales[list_index_for_sale],
             format(round(new_array_of_cumulative_sales[list_index_for_sale], 1), ',') + ' K',
             color='black', fontsize=10, fontweight='bold')
    plt.scatter(list_index_for_sale, new_array_of_cumulative_sales[list_index_for_sale], s=60, facecolors='#113d3c',
                edgecolors='white')

    plt.text(list_index_for_target - 1.3, cumulative_target_that_needs_to_plot[list_index_for_target],
             format(round(cumulative_target_that_needs_to_plot[list_index_for_target]), ',') + ' K',
             color='black', fontsize=10, fontweight='bold')
    plt.scatter(list_index_for_target, cumulative_target_that_needs_to_plot[list_index_for_target], s=60,
                facecolors='red', edgecolors='white')

    ax.set_ylim(ymin=0)
    ax.set_xlim(xmin=0)
    plt.xticks(np.arange(1, total_days + 1, 1), np.arange(1, total_days + 1, 1), fontsize=12)
    # plt.xlabel('Days', color='black', fontsize=14, fontweight='bold')
    plt.ylabel('Amount(K)', color='black', fontsize=10, fontweight='bold')
    plt.title('Cumulative Day Wise Stock & Sales', color='black', fontweight='bold', fontsize=12)

    plt.legend(['Target', 'Sales'])
    plt.yticks(np.arange(0, max(cumulative_target_that_needs_to_plot) + 0.4 * max(cumulative_target_that_needs_to_plot),
                         max(cumulative_target_that_needs_to_plot) / 5))
    plt.tight_layout()

    plt.savefig("./Images/Cumulative_Day_Wise_Target_vs_Sales.png")
    # plt.show()
    print('Cumulative day wise target sales generated')
