import pandas as pd;
from datetime import datetime, timedelta
import math

current_time = datetime(2022, 3, 4)


def load_and_clean_data(path):

    sheets = pd.read_excel(path, sheet_name=None)

    #taking both sheets
    orders_df =  sheets[list(sheets.keys())[0]]
    machine_df = sheets[list(sheets.keys())[1]]


    #orders
    orders_df.columns = [col.strip().lower().replace(" ", "_").replace(".", "").replace("(", "").replace(")", "") for col in orders_df.columns]

    #machines
    machine_df.columns = [col.strip().lower().replace(" ", "_").replace(".", "").replace("(", "").replace(")", "").replace('\n', '') for col in machine_df.columns]
    machine_df.rename(columns={"weekly_hours__100%": "weekly_hours"}, inplace=True)

    #cleaning order value
    orders_df["order_value"] = orders_df["order_value"].astype("string")
    orders_df['order_value'] = orders_df['order_value'].str.replace('.', '').str.replace(',', '.').astype(float)

    #converting date colums
    date_columns = ['needed_on', 'covering_demand_until', 'planned_start', 'planned_end']
    for col in date_columns:
        if col in orders_df.columns:
            orders_df[col] = pd.to_datetime(orders_df[col], errors='coerce')

    
    orders_df = orders_df[orders_df["processing_time"] > 0]
    orders_df = orders_df[orders_df["mould"].notnull()]
    orders_df = orders_df.dropna(subset=['machine', 'order_qty', 'processing_time', 'gross_processing_time'])

    #unit value
    orders_df["unit_value"] = orders_df["order_value"] / orders_df["order_qty"]

    #merge
    full_df = pd.merge(orders_df, machine_df, on="machine", how="left")
    return full_df


def calculate_ups(df):

    #1
    df["days_to_deadline"] = (df["needed_on"] - current_time).dt.days

    def fx(a):
        if a<= 1:
            return 50
        elif a<=3:
            return 40
        elif a<=7:
            return 30
        elif a<=14:
            return 20
        else:
            return 5
        
    df["urgency_score"] = df["days_to_deadline"].apply(fx)

    #2
    df["type_score"] = df["ref_sales_id"].notnull().astype(int) * 30

    #3
    max_order_value = df["order_value"].max()
    def gx(a):
        return (math.log1p(a)/ math.log1p(max_order_value)) * 10
    df["value_score"] = df["order_value"].apply(gx)

    #4
    max_priority = df["priority"].max()
    def bx(a):
        return (((max_priority) -a)/(max_priority) * 10)
    df["mrp_score"] = df["priority"].apply(bx)

    # Final ups calculation 
    df["ups"] = (df["urgency_score"] * .5) + (df["type_score"]*.3) + (df["value_score"]*0.1) + (df["mrp_score"]*0.1)

    df = df.sort_values(by="ups", ascending=False)
    return df


def create_mould_group(df):
    orders_by_mould = {}
    grouped_by_mould = df.groupby("mould")

    for mould, orders in grouped_by_mould:
        orders_by_mould[mould] = orders

    return orders_by_mould



def calculate_gross_processing_time(row, same_mould):
    
    if same_mould:
        availability = row["estimated_availability"]
        gross_time = (row["processing_time"]/availability)
    else:
        availability = row["estimated_availability"]
        # gross_time = row["gross_processing_time"]
        gross_time = ((row["net_processing_time_setup_&_processing"]/availability))

    return gross_time



def check_mould_constraints(order_c, mould_orders, current_scheduled_time):
    time_buffer_hours = 4
    value_threshold_percent=20

    c_gross_processing = calculate_gross_processing_time(order_c, True)
    estimated_completion = current_scheduled_time + timedelta(hours=c_gross_processing + time_buffer_hours)

    for index, higher_priority_order in mould_orders.iterrows():
        if(higher_priority_order["prod_order"] == order_c["prod_order"]):
            break

        #deadline
        if estimated_completion > higher_priority_order["needed_on"]:
            return False
        
        #check unit value
        order_c_unit_value = order_c["unit_value"]
        if higher_priority_order["unit_value"] >= order_c_unit_value:
            if higher_priority_order["unit_value"] >= order_c_unit_value * 1.2:
                return False


    mould_orders = mould_orders[mould_orders["prod_order"] != order_c["prod_order"]]
        

    return True

def find_opportunity_candidate(unscheduled_orders, current_mould, current_scheduled_time):

    candidate = None
    n = None
    for i, order in enumerate(unscheduled_orders):
        if order['mould'] == current_mould:
            n = i
            candidate = order
            break


    if candidate is None:
        return None
    
    candidate_gross_time = calculate_gross_processing_time(candidate, True)
    candidate_completion_time = current_scheduled_time + timedelta(hours=candidate_gross_time)

    
    for i in range(n):
        higher_order = unscheduled_orders[i]
        higher_order_gross_time = calculate_gross_processing_time(higher_order, False)

        higher_order_completion_time = candidate_completion_time + timedelta(hours=higher_order_gross_time)

        if higher_order_completion_time > higher_order["needed_on"]:
            return None
            
    return candidate


def sequencing_algo(machine_queue_df, orders_by_mould):


    unscheduled_orders = machine_queue_df.to_dict("records")
    final_schedules = []

    if(len(unscheduled_orders) < 1):
        return final_schedules

    current_order = unscheduled_orders[0]
    current_mould = None
    current_scheduled_time = current_time 


    while unscheduled_orders:
        
        current_order["planned_start"] = current_scheduled_time
        current_gross = calculate_gross_processing_time(current_order, current_order["mould"] == current_mould)
        current_order["gross_processing_time"] = current_gross
        current_order["planned_end"] = current_scheduled_time + timedelta(hours=current_gross)
        
        unscheduled_orders.remove(current_order)
        final_schedules.append(current_order)
        current_mould = current_order["mould"]

        orders_by_mould[current_mould] = orders_by_mould[current_mould][orders_by_mould[current_mould]["prod_order"] != current_order["prod_order"]]


        current_scheduled_time = current_order["planned_end"]
        

        if(len(unscheduled_orders) < 1):
            break

        order_c = find_opportunity_candidate(unscheduled_orders, current_mould, current_scheduled_time)

        if order_c:

            mould_orders = orders_by_mould[current_mould]
            if check_mould_constraints(order_c, mould_orders, current_scheduled_time):
                current_order = order_c
            else:
                current_order = unscheduled_orders[0]
        else:
            current_order = unscheduled_orders[0]

    return final_schedules




if __name__ == "__main__":
    path = "production planning.xlsx"
    full_df = load_and_clean_data(path)

    scored_df = calculate_ups(full_df)

    orders_by_mould = create_mould_group(scored_df)
    # print(orders_by_mould)

    unique_machines = scored_df["machine"].unique()

    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

    for machine_id in [350]:


        machine_queue_df = scored_df[scored_df["machine"] == machine_id]
        # print(machine_queue_df)

        schedule = sequencing_algo(machine_queue_df, orders_by_mould)
        # final_schedules[machine_id] = schedule

        schedule_df = pd.DataFrame(schedule)

        

        display_cols = [
            "prod_order", "machine","mould", "ups", "unit_value", "gross_processing_time",
            "needed_on", "planned_start", "planned_end", "value_score", "order_value"
        ]

        #without copy final_df is just a view of schedule_df
        final_df = schedule_df[display_cols].copy()

        final_df['ups'] = final_df['ups'].round(2)
        final_df['unit_value'] = final_df['unit_value'].round(2)
        final_df['gross_processing_time'] = final_df['gross_processing_time'].round(2)


        final_df["needed_on"] = final_df["needed_on"].dt.strftime('%Y-%m-%d')
        final_df["planned_start"] = final_df["planned_start"].dt.strftime('%Y-%m-%d %H:%M')
        final_df["planned_end"] = final_df["planned_end"].dt.strftime('%Y-%m-%d %H:%M')
        


        print(final_df)
        # print(orders_by_mould["W861902-00"])

        final_df.to_excel(writer, sheet_name=f'Machine_{machine_id}', index=False)


writer.close()
