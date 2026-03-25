from datetime import datetime, timedelta

import lunardate
import pandas as pd


# Convert Gregorian date to Lunar date
def solar_to_lunar(year, month, day):
    lunar_date = lunardate.LunarDate.fromSolarDate(year, month, day)
    return lunar_date.year, lunar_date.month, lunar_date.day


# Get the Earthly Branch corresponding to the hour
def get_earthly_branch(hour):
    branches = ["子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"]
    index = (hour + 1) // 2 % 12
    return branches[index], index + 1


def get_quotient_remainder(dividend, divisor):
    quotient, remainder = divmod(dividend, divisor)
    return quotient, remainder


# Calculate the divination result based on lunar date and Earthly Branch
def xiao_liuren(lunar_month, lunar_day, earthly_time):
    results = ["大安", "留连", "速喜", "赤口", "小吉", "空亡"]
    _, month_result = get_quotient_remainder(lunar_month, 6)
    _, day_result = get_quotient_remainder(lunar_day, 6)
    _, time_result = get_quotient_remainder(earthly_time, 6)
    total = month_result - 1 + day_result - 1 + time_result - 1
    _, div_result = get_quotient_remainder(total, 6)
    return results[div_result - 1]


# Main function to perform the divination
def xiao_liuren_divination(year, month, day, hour):
    _, lunar_month, lunar_day = solar_to_lunar(year, month, day)
    _, earthly_time = get_earthly_branch(hour)
    result = xiao_liuren(lunar_month, lunar_day, earthly_time)
    return result


def generate_datetime_range(start_datetime, end_datetime):
    current_datetime = start_datetime
    while current_datetime <= end_datetime:
        yield current_datetime
        current_datetime += timedelta(hours=1)


start_datetime = datetime(2024, 5, 1, 0, 0)
end_datetime = datetime(2026, 1, 1, 0, 0)

six_gods_meanings = {
    "大安": ["静", "不动", "吉利", "阳", "木"],
    "留连": ["慢", "反复", "拖延", "阴", "土"],
    "速喜": ["动", "快速", "吉利", "阳", "火"],
    "赤口": ["吵", "意见不和", "见金属利器伤害类", "阴", "金"],
    "小吉": ["动", "缓慢", "吉利", "阳", "水"],
    "空亡": ["无", "跟随性", "遇好则好，遇坏则坏", "阴", "土"],
}

# Create an empty list to store the results
results_list = []

# Generate dates and perform divination
for single_date in generate_datetime_range(start_datetime, end_datetime):
    try:
        result = xiao_liuren_divination(
            single_date.year, single_date.month, single_date.day, single_date.hour
        )
        # Calculate the meanings for the divination result
        meanings = six_gods_meanings.get(result, ["", "", "", "", ""])
        # Append the result along with the datetime components and meanings to the results list
        results_list.append(
            {
                "Year": single_date.year,
                "Month": single_date.month,
                "Day": single_date.day,
                "Hour": single_date.hour,
                "Result": result,
                "Meaning_1": meanings[0],
                "Meaning_2": meanings[1],
                "Meaning_3": meanings[2],
                "Meaning_4": meanings[3],
                "Meaning_5": meanings[4],
            }
        )
    except Exception as e:
        print(single_date, e)

# Create a DataFrame from the results list
results_df = pd.DataFrame(results_list)

# Save the DataFrame to a CSV file
results_df.to_csv("xiao_liuren_results.csv", index=False)
