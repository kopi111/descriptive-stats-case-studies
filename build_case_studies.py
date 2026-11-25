"""Build the five Descriptive Statistics case study submissions as Word files.

Student: Dwayne Aitken
Course: Probability and Statistics for AI

Every figure written in the answers is checked against Python's statistics
module before any file is produced, so a wrong answer stops the build.
"""

import statistics
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# The folder on the Desktop that receives the finished Word files.
OUTPUT_FOLDER = Path.home() / "Desktop" / "Descriptive Statistics Case Studies"

# The font and sizes used across every document.
BODY_FONT = "Times New Roman"
BODY_SIZE = Pt(12)
TABLE_SIZE = Pt(11)

# The grey used behind table header rows and label cells.
SHADE_COLOUR = "E7E6E6"

# The one source used: the lesson notes supplied for this unit.
REFERENCE_TEXT = (
    "Chithra. (n.d.). ",
    "Descriptive statistics: Mean, median and mode, measures of spread, and "
    "visualising data",
    " [Lesson notes]. Amber Heart Coding Academy.",
)

MEAN_FORMULA = "Mean = Sum of all values ÷ Number of values"
RANGE_FORMULA = "Range = Maximum value − Minimum value"
VARIANCE_FORMULA = "Population variance = Σ(x − Mean)² ÷ N"
MEDIAN_ODD_RULE = "Median = the middle value of the ordered data"
MEDIAN_EVEN_RULE = "Median = (first middle value + second middle value) ÷ 2"


def labelled_rows(prefix, count):
    """Return row labels such as Patient 1, Patient 2 and so on."""
    return [f"{prefix} {number}" for number in range(1, count + 1)]


def display_value(value):
    """Return a value as printed text, with thousands commas and true minus signs."""
    if isinstance(value, (int, float)):
        return f"{value:,}".replace("-", "−")
    return str(value)


def joined(values):
    """Return the values as a comma-separated line of text."""
    return ", ".join(display_value(value) for value in values)


def sum_line(values):
    """Return the addition written out, ending with its total."""
    additions = " + ".join(display_value(value) for value in values)
    return f"{additions} = {display_value(sum(values))}"


def frequency_rows(values):
    """Return (value, frequency) rows in a stable, readable order."""
    counts = Counter(values)
    return [(value, counts[value]) for value in sorted(counts, key=str)]


def require_equal(actual, expected, description):
    """Stop the build when a written answer disagrees with the calculation."""
    if round(actual, 2) != expected:
        raise ValueError(f"{description}: calculated {actual}, written {expected}")


def check_mean(case):
    """Confirm the total and the mean written for one mean case study."""
    values = case["values"]
    require_equal(sum(values), case["total"], f"{case['title']} total")
    require_equal(statistics.mean(values), case["answer"], f"{case['title']} mean")


def check_median(case):
    """Confirm the median written for one median case study."""
    require_equal(
        statistics.median(case["values"]),
        case["answer"],
        f"{case['title']} median",
    )


def check_mode(case):
    """Confirm the mode and its frequency for one mode case study."""
    counts = Counter(case["values"])
    if statistics.mode(case["values"]) != case["answer"]:
        raise ValueError(f"{case['title']}: mode is not {case['answer']}")
    if counts[case["answer"]] != case["highest_frequency"]:
        raise ValueError(f"{case['title']}: wrong highest frequency")
    if list(counts.values()).count(case["highest_frequency"]) != 1:
        raise ValueError(f"{case['title']}: more than one mode")


def check_range(case):
    """Confirm the range written for one range case study."""
    values = case["values"]
    require_equal(max(values) - min(values), case["answer"], f"{case['title']} range")


def check_variance(case):
    """Confirm the mean, the squared total and the variance for one case study."""
    values = case["values"]
    mean = statistics.mean(values)
    squared_total = sum((value - mean) ** 2 for value in values)
    require_equal(mean, case["mean"], f"{case['title']} mean")
    require_equal(squared_total, case["squared_total"], f"{case['title']} squares")
    require_equal(
        statistics.pvariance(values),
        case["answer"],
        f"{case['title']} variance",
    )
    require_equal(
        statistics.pstdev(values),
        case["standard_deviation"],
        f"{case['title']} standard deviation",
    )


MEAN_CASES = [
    {
        "title": "Case Study 1: Vegetable Market Sales",
        "brief": "A vegetable vendor recorded the number of kilograms of tomatoes "
        "sold over six days.",
        "headers": ("Day", "Quantity sold (kg)"),
        "labels": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
        "values": [12, 15, 18, 10, 20, 15],
        "total": 90,
        "answer": 15,
        "calculation": "90 ÷ 6 = 15",
        "final": "15 kg per day",
        "interpretation": "The vendor sold an average of 15 kg of tomatoes per day "
        "over the six days.",
        "tasks": [
            ("Calculate the total quantity sold.", "The total quantity sold is 90 kg."),
            (
                "Find the mean quantity sold per day.",
                "The mean quantity sold is 90 ÷ 6 = 15 kg per day.",
            ),
            (
                "Explain what the mean tells the vendor.",
                "The mean tells the vendor that a typical day over this period "
                "produced sales of about 15 kg. The vendor can use this figure as a "
                "starting point for daily stock orders. For example, ordering close "
                "to 15 kg supports most days, although Wednesday and Friday sales "
                "exceeded this amount.",
            ),
        ],
    },
    {
        "title": "Case Study 2: Student Academic Performance",
        "brief": "A student obtained the following marks in six AI and Mathematics "
        "assessments.",
        "headers": ("Assessment", "Marks"),
        "labels": labelled_rows("Test", 6),
        "values": [60, 65, 70, 75, 80, 90],
        "total": 440,
        "answer": 73.33,
        "calculation": "440 ÷ 6 = 73.33 (rounded to two decimal places)",
        "final": "73.33 marks",
        "interpretation": "The student averaged 73.33 marks across the six assessments.",
        "tasks": [
            ("Calculate the total marks.", "The total is 440 marks."),
            (
                "Calculate the mean mark.",
                "The mean mark is 440 ÷ 6 = 73.33, rounded to two decimal places.",
            ),
            (
                "Is the mean greater than 70?",
                "Yes. The mean of 73.33 is 3.33 marks greater than 70.",
            ),
            (
                "Explain how the mean can help a teacher understand student "
                "performance.",
                "The mean gives the teacher one figure that summarizes overall "
                "performance across several assessments. The teacher can compare it "
                "with a pass mark, a class average or the student's earlier results. "
                "For example, a mean of 73.33 shows solid overall performance, while "
                "the individual marks show steady improvement from 60 to 90 that the "
                "mean alone does not reveal.",
            ),
        ],
    },
    {
        "title": "Case Study 3: Hospital Patient Waiting Time",
        "brief": "A hospital recorded the waiting time of five patients before "
        "consultation.",
        "headers": ("Patient", "Waiting time (minutes)"),
        "labels": labelled_rows("Patient", 5),
        "values": [10, 15, 20, 25, 30],
        "total": 100,
        "answer": 20,
        "calculation": "100 ÷ 5 = 20",
        "final": "20 minutes",
        "interpretation": "These patients waited an average of 20 minutes before "
        "consultation.",
        "tasks": [
            ("Find the total waiting time.", "The total waiting time is 100 minutes."),
            (
                "Calculate the mean waiting time.",
                "The mean waiting time is 100 ÷ 5 = 20 minutes.",
            ),
            (
                "What does the mean tell the hospital management?",
                "The mean tells management that a patient in this group waited about "
                "20 minutes on average. Management can compare this figure with a "
                "service target to decide whether staffing or scheduling needs "
                "attention. Because only five patients are included, the result "
                "should be confirmed with more records before a decision is made.",
            ),
            (
                "How can an AI-based hospital management system use average waiting "
                "time?",
                "An AI-based system can track the average waiting time by hour, day "
                "and department to identify busy periods. It can then support staff "
                "scheduling and give patients an estimated wait. For example, the "
                "system could alert a supervisor when the current average rises well "
                "above the usual 20 minutes.",
            ),
        ],
    },
    {
        "title": "Case Study 4: Supermarket Customer Spending",
        "brief": "A supermarket recorded the amount spent by six customers.",
        "headers": ("Customer", "Amount spent (₹)"),
        "labels": labelled_rows("Customer", 6),
        "values": [200, 300, 250, 400, 350, 300],
        "total": 1800,
        "answer": 300,
        "calculation": "1,800 ÷ 6 = 300",
        "final": "₹300 per customer",
        "interpretation": "Each of these customers spent an average of ₹300.",
        "tasks": [
            (
                "Calculate the total customer spending.",
                "The total customer spending is ₹1,800.",
            ),
            (
                "Find the mean spending per customer.",
                "The mean spending is ₹1,800 ÷ 6 = ₹300 per customer.",
            ),
            (
                "What does the mean tell the supermarket?",
                "The mean tells the supermarket that the average purchase in this "
                "group is ₹300. The supermarket can use this figure to estimate "
                "revenue from the number of customers served. For example, 100 "
                "customers at this average would suggest revenue of about ₹30,000.",
            ),
            (
                "How can an AI-based retail system use this information?",
                "An AI-based retail system can use average spending as an input when "
                "forecasting sales and grouping customers. It can compare each "
                "customer with the ₹300 average to support relevant promotions. For "
                "example, a customer spending close to ₹300 could be shown a bundle "
                "offer at a slightly higher value.",
            ),
        ],
    },
    {
        "title": "Case Study 5: Weather Monitoring",
        "brief": "A weather station recorded the following temperatures over seven "
        "days.",
        "headers": ("Day", "Temperature (°C)"),
        "labels": [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ],
        "values": [28, 30, 31, 29, 32, 34, 28],
        "total": 212,
        "answer": 30.29,
        "calculation": "212 ÷ 7 = 30.29 (rounded to two decimal places)",
        "final": "30.29°C",
        "interpretation": "The average temperature over the seven days was 30.29°C.",
        "tasks": [
            (
                "Calculate the total temperature values.",
                "The total of the temperature values is 212.",
            ),
            (
                "Find the mean temperature.",
                "The mean temperature is 212 ÷ 7 = 30.29°C, rounded to two decimal "
                "places.",
            ),
            (
                "Which days recorded temperatures above the mean?",
                "Wednesday (31°C), Friday (32°C) and Saturday (34°C) recorded "
                "temperatures above the mean. Tuesday, at 30°C, falls just below it.",
            ),
            (
                "Explain how average temperature can be useful in weather prediction.",
                "Average temperature provides a baseline for what is normal during a "
                "period. Prediction models can compare new readings with this "
                "baseline to detect warming or cooling trends. For example, a reading "
                "several degrees above 30.29°C could indicate that an unusually hot "
                "period is developing.",
            ),
        ],
    },
]

MEDIAN_CASES = [
    {
        "title": "Case Study 1: Vegetable Market Sales",
        "brief": "A vegetable vendor recorded the number of kilograms of tomatoes "
        "sold over seven days.",
        "headers": ("Day", "Quantity sold (kg)"),
        "labels": [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ],
        "values": [15, 20, 12, 25, 18, 30, 22],
        "answer": 20,
        "middle": "The 4th value, which is 20",
        "calculation": "n = 7 is odd, so the median is the single middle value: 20",
        "final": "20 kg",
        "interpretation": "On half of the days the vendor sold 20 kg or less, and on "
        "half of the days the vendor sold 20 kg or more.",
        "tasks": [
            (
                "Arrange the sales quantities in ascending order.",
                "12, 15, 18, 20, 22, 25, 30",
            ),
            ("Count the number of observations.", "There are 7 observations."),
            (
                "Identify the middle value.",
                "The middle value is the 4th value in the ordered list, which is 20.",
            ),
            ("Calculate the median sales quantity.", "The median is 20 kg."),
            (
                "Explain what the median tells the vendor.",
                "The median tells the vendor that 20 kg is a typical day's sales, "
                "with an equal number of days falling above and below it. Unlike the "
                "mean, the figure is not pulled upward by the strong 30 kg day. The "
                "vendor can therefore use 20 kg as a dependable guide for planning "
                "normal daily stock.",
            ),
        ],
    },
    {
        "title": "Case Study 2: Student Academic Performance",
        "brief": "An AI student received the following marks in eight assessments.",
        "headers": ("Assessment", "Marks"),
        "labels": labelled_rows("Test", 8),
        "values": [45, 60, 75, 55, 80, 70, 90, 65],
        "answer": 67.5,
        "middle": "The 4th and 5th values, which are 65 and 70",
        "calculation": "(65 + 70) ÷ 2 = 135 ÷ 2 = 67.5",
        "final": "67.5 marks",
        "interpretation": "The student scored below 67.5 on half of the assessments "
        "and above 67.5 on the other half.",
        "tasks": [
            (
                "Arrange the marks in ascending order.",
                "45, 55, 60, 65, 70, 75, 80, 90",
            ),
            ("Identify the two middle values.", "The two middle values are 65 and 70."),
            ("Calculate the median mark.", "The median is (65 + 70) ÷ 2 = 67.5 marks."),
            (
                "Explain how median can help a teacher understand student performance.",
                "The median shows the teacher the student's typical level of "
                "performance without letting one very high or very low score distort "
                "it. For example, the low mark of 45 and the high mark of 90 do not "
                "change the median as long as they remain at the ends of the ordered "
                "list. This supports a fairer view of the student's usual standard.",
            ),
        ],
    },
    {
        "title": "Case Study 3: Hospital Patient Waiting Time",
        "brief": "A hospital recorded the waiting time of seven patients before "
        "consultation.",
        "headers": ("Patient", "Waiting time (minutes)"),
        "labels": labelled_rows("Patient", 7),
        "values": [10, 15, 20, 25, 30, 35, 90],
        "answer": 25,
        "middle": "The 4th value, which is 25",
        "calculation": "n = 7 is odd, so the median is the single middle value: 25",
        "final": "25 minutes",
        "interpretation": "Half of these patients waited 25 minutes or less, and half "
        "waited 25 minutes or more.",
        "tasks": [
            (
                "Arrange the waiting times in ascending order.",
                "10, 15, 20, 25, 30, 35, 90",
            ),
            ("Find the median waiting time.", "The median waiting time is 25 minutes."),
            (
                "Why does the 90-minute waiting time not become the median?",
                "The median depends on position in the ordered list and not on the "
                "size of the values. The 90-minute observation is the largest value, "
                "so it sits at the end of the list as the 7th value, while the median "
                "is the 4th value. Changing 90 to any other value above 25 would "
                "leave the median at 25 minutes.",
            ),
            (
                "How can an AI-based hospital management system use median waiting "
                "time?",
                "An AI-based system can report the median as the typical patient "
                "experience because one extreme delay does not distort it. In this "
                "data the mean is about 32.14 minutes, while the median is 25 "
                "minutes. The system can then track the median by department and "
                "flag extreme cases such as the 90-minute wait for separate review.",
            ),
        ],
    },
    {
        "title": "Case Study 4: Supermarket Customer Spending",
        "brief": "A supermarket recorded the amounts spent by six customers.",
        "headers": ("Customer", "Amount spent (₹)"),
        "labels": labelled_rows("Customer", 6),
        "values": [100, 200, 150, 1000, 250, 300],
        "answer": 225,
        "middle": "The 3rd and 4th values, which are 200 and 250",
        "calculation": "(200 + 250) ÷ 2 = 450 ÷ 2 = 225",
        "final": "₹225",
        "interpretation": "Half of these customers spent less than ₹225, and half "
        "spent more than ₹225.",
        "tasks": [
            (
                "Arrange the spending amounts in ascending order.",
                "₹100, ₹150, ₹200, ₹250, ₹300, ₹1,000",
            ),
            (
                "Identify the two middle values.",
                "The two middle values are ₹200 and ₹250.",
            ),
            (
                "Calculate the median spending.",
                "The median spending is (200 + 250) ÷ 2 = ₹225.",
            ),
            (
                "Why is median useful when one customer spends much more than the "
                "others?",
                "The median is not pulled upward by one unusually large purchase. In "
                "this data the mean is about ₹333.33, which is more than five of the "
                "six customers spent, because of the ₹1,000 purchase. The median of "
                "₹225 therefore gives a more realistic picture of a typical customer.",
            ),
            (
                "How can AI-based retail analytics use this information?",
                "AI-based retail analytics can compare the median with the mean to "
                "detect when a few high spenders are inflating the average. The "
                "system can then treat typical customers and high-value customers as "
                "separate groups. For example, general promotions could be designed "
                "around ₹225, while the ₹1,000 customer receives a separate offer.",
            ),
        ],
    },
    {
        "title": "Case Study 5: Weather Monitoring",
        "brief": "A weather station recorded the following temperatures over eight "
        "days.",
        "headers": ("Day", "Temperature (°C)"),
        "labels": [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
            "Monday",
        ],
        "values": [28, 30, 32, 29, 35, 31, 33, 40],
        "answer": 31.5,
        "middle": "The 4th and 5th values, which are 31 and 32",
        "calculation": "(31 + 32) ÷ 2 = 63 ÷ 2 = 31.5",
        "final": "31.5°C",
        "interpretation": "Half of the recorded temperatures fall below 31.5°C, and "
        "half fall above it.",
        "tasks": [
            (
                "Arrange the temperatures in ascending order.",
                "28, 29, 30, 31, 32, 33, 35, 40",
            ),
            ("Identify the two middle values.", "The two middle values are 31 and 32."),
            (
                "Calculate the median temperature.",
                "The median temperature is (31 + 32) ÷ 2 = 31.5°C.",
            ),
            (
                "Explain why median is useful when unusual temperatures occur.",
                "The median remains stable when one reading is unusually high or "
                "low. The 40°C reading raises the mean to 32.25°C, while the median "
                "stays at 31.5°C. Without that reading, the mean would be about "
                "31.14°C and the median 31°C, so the median moves less. It therefore "
                "gives a clearer view of normal conditions.",
            ),
            (
                "How can weather prediction systems use median temperature?",
                "Weather prediction systems can use the median as a stable baseline "
                "for normal conditions in a period. New readings can be compared with "
                "this baseline to identify unusual events or possible sensor errors. "
                "For example, the 40°C reading could be flagged for checking because "
                "it is well above the 31.5°C median.",
            ),
        ],
    },
]

MODE_CASES = [
    {
        "title": "Case Study 1: Vegetable Market Sales",
        "brief": "A vegetable vendor recorded the number of kilograms of tomatoes "
        "sold over seven days.",
        "headers": ("Day", "Quantity sold (kg)"),
        "labels": [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ],
        "values": [10, 15, 10, 20, 15, 10, 25],
        "value_heading": "Quantity sold (kg)",
        "answer": 10,
        "highest_frequency": 3,
        "final": "10 kg",
        "interpretation": "The most common daily sales quantity was 10 kg, which "
        "occurred on three of the seven days.",
        "tasks": [
            (
                "Prepare a frequency table for the quantities sold.",
                "The frequency table appears above.",
            ),
            ("Identify the highest frequency.", "The highest frequency is 3."),
            ("Find the mode.", "The mode is 10 kg."),
            (
                "Explain which quantity was sold most frequently.",
                "The quantity sold most frequently was 10 kg, which occurred on "
                "Monday, Wednesday and Saturday.",
            ),
            (
                "How can the vendor use this information to plan stock?",
                "The vendor can treat 10 kg as the minimum amount to have available "
                "each day because this level of sales occurs most often. The vendor "
                "should still hold some extra stock, since sales reached 15 kg or "
                "more on the other four days. This approach balances the risk of "
                "waste with the risk of running short.",
            ),
        ],
    },
    {
        "title": "Case Study 2: Student Academic Performance",
        "brief": "A teacher recorded the examination marks of ten students in an AI "
        "subject.",
        "headers": ("Student", "Marks"),
        "labels": labelled_rows("Student", 10),
        "values": [60, 70, 80, 70, 90, 60, 70, 80, 70, 50],
        "value_heading": "Marks",
        "answer": 70,
        "highest_frequency": 4,
        "final": "70 marks",
        "interpretation": "The most common mark was 70, which was obtained by four of "
        "the ten students.",
        "tasks": [
            ("Prepare a frequency table.", "The frequency table appears above."),
            (
                "Identify the mark that occurs most frequently.",
                "The mark of 70 occurs most frequently, with a frequency of 4.",
            ),
            ("Calculate the mode.", "The mode is 70 marks."),
            (
                "Explain what the most frequently obtained mark tells the teacher.",
                "The mode tells the teacher that 70 is the score most commonly "
                "achieved in this class. It shows where the largest group of students "
                "is performing. However, it does not describe the other six students, "
                "whose marks range from 50 to 90, so the teacher should review the "
                "full distribution as well.",
            ),
            (
                "How can educational analytics use this information?",
                "Educational analytics can use the most common mark to identify the "
                "level reached by the largest group of students. It can compare modes "
                "across subjects or terms to reveal changes over time. For example, "
                "if the mode moved from 70 to 60 in a later term, the school could "
                "review teaching support for that subject.",
            ),
        ],
    },
    {
        "title": "Case Study 3: Hospital Patient Blood Groups",
        "brief": "A hospital recorded the blood groups of eight hypothetical patients.",
        "headers": ("Patient", "Blood group"),
        "labels": labelled_rows("Patient", 8),
        "values": ["A", "O", "B", "O", "AB", "O", "A", "O"],
        "value_heading": "Blood group",
        "answer": "O",
        "highest_frequency": 4,
        "final": "Blood group O",
        "interpretation": "Blood group O was the most common in this group, occurring "
        "in four of the eight hypothetical patients.",
        "tasks": [
            (
                "Prepare a frequency table for the blood groups.",
                "The frequency table appears above.",
            ),
            (
                "Identify the most frequently occurring blood group.",
                "Blood group O occurs most frequently, with a frequency of 4.",
            ),
            ("Find the mode.", "The mode is blood group O."),
            (
                "Explain how frequency analysis can summarise hospital records.",
                "Frequency analysis reduces a long list of individual records to a "
                "short table of categories and counts. Staff can then see the pattern "
                "at a glance without reading each record. For example, these eight "
                "records become four blood groups with counts of 2, 1, 1 and 4, which "
                "supports easier reporting and planning.",
            ),
            (
                "Why should this small sample not be used to conclude the most common "
                "blood group in the entire population?",
                "Eight hypothetical patients cannot represent an entire population. "
                "The sample is very small, so a few different records could change "
                "the result, and these patients were not selected to reflect the "
                "wider population. A reliable conclusion would require a much larger "
                "and representative sample of real records.",
            ),
        ],
    },
    {
        "title": "Case Study 4: Supermarket Customer Purchases",
        "brief": "An online supermarket recorded the following product categories "
        "purchased by ten customers.",
        "headers": ("Customer", "Product category"),
        "labels": labelled_rows("Customer", 10),
        "values": [
            "Clothing",
            "Food",
            "Clothing",
            "Electronics",
            "Food",
            "Clothing",
            "Books",
            "Clothing",
            "Food",
            "Clothing",
        ],
        "value_heading": "Product category",
        "answer": "Clothing",
        "highest_frequency": 5,
        "final": "Clothing",
        "interpretation": "Clothing was the most popular category, purchased by five "
        "of the ten customers.",
        "tasks": [
            ("Prepare a frequency table.", "The frequency table appears above."),
            (
                "Identify the category with the highest frequency.",
                "Clothing has the highest frequency, at 5.",
            ),
            ("Find the mode.", "The mode is Clothing."),
            (
                "Explain which product category is most popular in this sample.",
                "Clothing is the most popular category in this sample because half "
                "of the customers purchased it. Food follows with 3 purchases, while "
                "Electronics and Books have 1 purchase each.",
            ),
            (
                "How can an AI-based shopping system use this information for "
                "inventory planning?",
                "An AI-based shopping system can use category frequencies to forecast "
                "demand and allocate stock. Popular categories such as Clothing would "
                "receive higher stock levels, while slower categories would hold less "
                "to reduce unsold items. Because this sample has only ten customers, "
                "the system should confirm the pattern with a larger purchase history.",
            ),
        ],
    },
    {
        "title": "Case Study 5: Weather Monitoring",
        "brief": "A weather station recorded the following temperatures over ten days.",
        "headers": ("Day", "Temperature (°C)"),
        "labels": [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
            "Monday",
            "Tuesday",
            "Wednesday",
        ],
        "values": [30, 32, 30, 31, 30, 33, 32, 30, 31, 30],
        "value_heading": "Temperature (°C)",
        "answer": 30,
        "highest_frequency": 5,
        "final": "30°C",
        "interpretation": "The most common temperature was 30°C, which was recorded on "
        "five of the ten days.",
        "tasks": [
            ("Prepare a frequency table.", "The frequency table appears above."),
            (
                "Identify the most frequently recorded temperature.",
                "The most frequently recorded temperature is 30°C, with a frequency "
                "of 5.",
            ),
            ("Calculate the mode.", "The mode is 30°C."),
            (
                "Explain what the mode tells us about the temperature records.",
                "The mode shows that 30°C was the usual reading during this period, "
                "occurring on half of the days. All other readings were only 1 to "
                "3 degrees higher, and no reading fell below 30°C. The records "
                "therefore suggest fairly steady conditions.",
            ),
            (
                "How can weather analytics use frequency information?",
                "Weather analytics can use frequency information to show which "
                "temperatures are common and which are rare at a location. This "
                "supports the detection of unusual readings and long-term changes. "
                "For example, if 33°C became the most frequent reading in later "
                "periods, analysts could investigate a possible warming pattern.",
            ),
        ],
    },
]

RANGE_CASES = [
    {
        "title": "Case Study 1: Vegetable Market",
        "brief": "A vegetable vendor records the number of tomato baskets sold over "
        "five days.",
        "values": [12, 18, 15, 25, 30],
        "data_text": "12, 18, 15, 25, 30",
        "answer": 18,
        "calculation": "30 − 12 = 18 baskets",
        "tasks": [
            (
                "Identify the highest and lowest sales.",
                "The highest sales figure is 30 baskets, and the lowest is 12 baskets.",
            ),
            ("Calculate the range.", "Range = 30 − 12 = 18 baskets."),
            (
                "Explain what the range tells the vendor.",
                "The range tells the vendor that daily sales differed by as much as "
                "18 baskets over the five days. This is a fairly wide difference, so "
                "demand was not steady. The vendor can use this information to avoid "
                "fixing stock at a single level, although the range does not show how "
                "the three other days were distributed.",
            ),
        ],
    },
    {
        "title": "Case Study 2: Student Attendance",
        "brief": "The attendance percentages of five students are as follows.",
        "values": [65, 75, 80, 90, 95],
        "data_text": "65, 75, 80, 90, 95",
        "answer": 30,
        "calculation": "95 − 65 = 30 percentage points",
        "tasks": [
            (
                "Find the maximum and minimum attendance.",
                "The maximum attendance is 95%, and the minimum is 65%.",
            ),
            ("Calculate the range.", "Range = 95 − 65 = 30 percentage points."),
            (
                "Explain whether the students have similar attendance percentages.",
                "The students do not have very similar attendance percentages. A "
                "range of 30 percentage points is wide, with one student attending "
                "65% of the time and another attending 95%. The other values also "
                "spread across this gap instead of clustering together. The school "
                "may therefore wish to follow up with the students at the lower end.",
            ),
        ],
    },
    {
        "title": "Case Study 3: Hospital Waiting Time",
        "brief": "Five patients wait for the following number of minutes.",
        "values": [10, 15, 20, 25, 60],
        "data_text": "10, 15, 20, 25, 60",
        "answer": 50,
        "calculation": "60 − 10 = 50 minutes",
        "tasks": [
            (
                "Calculate the waiting-time range.",
                "Range = 60 − 10 = 50 minutes.",
            ),
            (
                "Identify the longest waiting time.",
                "The longest waiting time is 60 minutes.",
            ),
            (
                "Explain why the hospital should investigate the 60-minute observation.",
                "The 60-minute observation is far above the other four waiting "
                "times, which are all 25 minutes or less. It is 35 minutes longer "
                "than the next value, and it accounts for most of the 50-minute "
                "range, since the range would be only 15 minutes without it. The "
                "hospital should determine whether it reflects a staffing delay, an "
                "emergency or a recording error, which supports accountability and "
                "better patient service.",
            ),
        ],
    },
    {
        "title": "Case Study 4: Online Shopping",
        "brief": "Five customers spend the following amounts.",
        "values": [250, 400, 500, 700, 1000],
        "data_text": "₹250, ₹400, ₹500, ₹700, ₹1,000",
        "answer": 750,
        "calculation": "₹1,000 − ₹250 = ₹750",
        "tasks": [
            ("Calculate the spending range.", "Range = ₹1,000 − ₹250 = ₹750."),
            (
                "Explain what it tells the shop owner.",
                "The range tells the shop owner that the highest and lowest customer "
                "spending differ by ₹750. Spending therefore varies widely, which "
                "suggests that the shop serves both smaller and larger buyers. The "
                "owner can consider offers that suit both groups.",
            ),
            (
                "Suggest one other measure that can describe customer spending.",
                "The mean is one other useful measure. The mean spending is "
                "₹2,850 ÷ 5 = ₹570 per customer. While the range describes only the "
                "gap between the two extreme values, the mean uses every customer's "
                "spending and shows the average purchase.",
            ),
        ],
    },
    {
        "title": "Case Study 5: Weather Monitoring",
        "brief": "A weather station records these temperatures over five days.",
        "values": [26, 28, 30, 32, 36],
        "data_text": "26, 28, 30, 32, 36°C",
        "answer": 10,
        "calculation": "36 − 26 = 10°C",
        "tasks": [
            ("Calculate the temperature range.", "Range = 36 − 26 = 10°C."),
            (
                "Identify the hottest and coolest days using the recorded sequence.",
                "Using the recorded sequence, Day 5 is the hottest day at 36°C, and "
                "Day 1 is the coolest day at 26°C.",
            ),
            (
                "Explain how this information can help weather monitoring.",
                "The range shows how much the temperature changed during the period. "
                "In this sequence, the temperature rose every day and finished 10°C "
                "above its starting point. A weather station can compare this range "
                "with other weeks to identify unusual changes. For example, a much "
                "wider range could prompt further analysis or an early heat advisory.",
            ),
        ],
    },
]

VARIANCE_CASES = [
    {
        "title": "Case Study 1: Vegetable Sales",
        "brief": "A vegetable seller records the number of baskets sold over five days.",
        "value_heading": "Baskets sold (x)",
        "unit": "baskets",
        "values": [10, 12, 14, 16, 18],
        "mean": 14,
        "squared_total": 40,
        "answer": 8,
        "standard_deviation": 2.83,
        "final": "8 baskets²",
        "tasks": [
            (
                "Calculate the population variance.",
                "The population variance is 40 ÷ 5 = 8 squared baskets.",
            ),
            (
                "Explain whether the daily sales are close to their average.",
                "The daily sales are fairly close to their average of 14 baskets. "
                "No day differs from the average by more than 4 baskets, and the "
                "variance of 8 corresponds to a standard deviation of about 2.83 "
                "baskets. Sales are therefore reasonably consistent, which supports "
                "dependable stock planning.",
            ),
        ],
    },
    {
        "title": "Case Study 2: Student Performance",
        "brief": "Five students obtain the following marks.",
        "value_heading": "Marks (x)",
        "unit": "marks",
        "values": [40, 50, 60, 70, 80],
        "mean": 60,
        "squared_total": 1000,
        "answer": 200,
        "standard_deviation": 14.14,
        "final": "Mean = 60 marks; variance = 200 marks²",
        "tasks": [
            ("Calculate the mean.", "The mean is 300 ÷ 5 = 60 marks."),
            (
                "Calculate the population variance.",
                "The population variance is 1,000 ÷ 5 = 200 squared marks.",
            ),
            (
                "Explain what the variance tells the teacher about the marks.",
                "The variance tells the teacher that the marks are widely spread "
                "around the mean of 60. A variance of 200 corresponds to a standard "
                "deviation of about 14.14 marks, and the marks lie up to 20 marks "
                "from the mean. The class therefore contains quite different levels "
                "of performance, so the teacher may need to provide extra support "
                "for some students and further challenge for others.",
            ),
        ],
    },
    {
        "title": "Case Study 3: Hospital Waiting Time",
        "brief": "The waiting times of five patients are as follows.",
        "value_heading": "Waiting time in minutes (x)",
        "unit": "minutes",
        "values": [5, 10, 15, 20, 25],
        "mean": 15,
        "squared_total": 250,
        "answer": 50,
        "standard_deviation": 7.07,
        "final": "50 minutes²",
        "tasks": [
            (
                "Calculate the population variance.",
                "The population variance is 250 ÷ 5 = 50 squared minutes.",
            ),
            (
                "Explain how waiting-time variation can be useful in hospital service "
                "analysis.",
                "Waiting-time variation shows how consistent the service is, which "
                "an average alone cannot reveal. Here the mean is 15 minutes, and a "
                "variance of 50 corresponds to a standard deviation of about 7.07 "
                "minutes. For example, two clinics could share a 15-minute mean, "
                "while the clinic with the higher variance gives patients less "
                "predictable waits. Management can use this information to review "
                "scheduling and maintain a more dependable service.",
            ),
        ],
    },
    {
        "title": "Case Study 4: Customer Spending",
        "brief": "Five customers spend the following amounts.",
        "value_heading": "Amount spent in ₹ (x)",
        "unit": "₹",
        "values": [100, 150, 200, 250, 300],
        "mean": 200,
        "squared_total": 25000,
        "answer": 5000,
        "standard_deviation": 70.71,
        "final": "5,000 ₹²",
        "tasks": [
            (
                "Calculate the population variance.",
                "The population variance is 25,000 ÷ 5 = 5,000 squared rupees.",
            ),
            (
                "Explain what the result tells the shop owner about customer spending.",
                "The result tells the shop owner that customer spending varies "
                "considerably around the mean of ₹200. A variance of 5,000 "
                "corresponds to a standard deviation of about ₹70.71, and spending "
                "ranges from ₹100 to ₹300. The owner should therefore not assume "
                "that every customer spends about the same amount, and can plan "
                "offers for both lower and higher spenders.",
            ),
        ],
    },
    {
        "title": "Case Study 5: Weather Monitoring",
        "brief": "A weather station records these temperatures.",
        "value_heading": "Temperature in °C (x)",
        "unit": "°C",
        "values": [26, 28, 30, 32, 34],
        "mean": 30,
        "squared_total": 40,
        "answer": 8,
        "standard_deviation": 2.83,
        "final": "8 °C²",
        "tasks": [
            (
                "Calculate the population variance.",
                "The population variance is 40 ÷ 5 = 8 squared degrees Celsius.",
            ),
            (
                "Explain whether the recorded temperatures show a small or large "
                "spread.",
                "The recorded temperatures show a small spread. A variance of 8 "
                "corresponds to a standard deviation of about 2.83°C, and no reading "
                "lies more than 4°C from the mean of 30°C. The temperatures therefore "
                "remained fairly stable during the recorded period.",
            ),
        ],
    },
]


def set_cell_shade(cell):
    """Shade one table cell so headers and labels stand out when printed."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), SHADE_COLOUR)
    cell._tc.get_or_add_tcPr().append(shading)


def keep_row_together(row):
    """Stop a table row from splitting across two pages."""
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def write_cell(cell, text, bold=False, keep_with_next=True):
    """Write text into a table cell using the shared table formatting."""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.keep_with_next = keep_with_next
    run = paragraph.add_run(display_value(text))
    run.font.size = TABLE_SIZE
    run.bold = bold


def add_table(document, rows, widths, header=True, label_column=False):
    """Add a bordered table, keeping all of its rows on one page."""
    table = document.add_table(rows=len(rows), cols=len(widths))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for column, width in zip(table.columns, widths):
        column.width = Inches(width)
    last_row_index = len(rows) - 1
    for row_index, row_values in enumerate(rows):
        keep_row_together(table.rows[row_index])
        for column_index, value in enumerate(row_values):
            cell = table.cell(row_index, column_index)
            cell.width = Inches(widths[column_index])
            is_header = header and row_index == 0
            is_label = label_column and column_index == 0
            write_cell(
                cell,
                value,
                bold=is_header or is_label,
                keep_with_next=row_index != last_row_index,
            )
            if is_header or is_label:
                set_cell_shade(cell)
    add_gap(document)
    return table


def add_gap(document):
    """Add a small space after a table."""
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 0.6


def add_page_number_field(paragraph):
    """Insert a live page number that Word and PDF conversion both fill in."""
    run = paragraph.add_run()
    for field_type, instruction in (("begin", None), (None, "PAGE"), ("end", None)):
        if instruction:
            element = OxmlElement("w:instrText")
            element.text = instruction
        else:
            element = OxmlElement("w:fldChar")
            element.set(qn("w:fldCharType"), field_type)
        run._r.append(element)


def style_heading(document, style_name, size, before, after):
    """Give a built-in heading style the paper's font, size and spacing."""
    style = document.styles[style_name]
    style.font.name = BODY_FONT
    style.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    style.element.rPr.rFonts.set(qn("w:asciiTheme"), BODY_FONT)
    style.element.rPr.rFonts.set(qn("w:hAnsiTheme"), BODY_FONT)
    style.font.size = size
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = before
    style.paragraph_format.space_after = after
    style.paragraph_format.keep_with_next = True


def new_document(topic):
    """Create a Letter-sized document with shared styles and a page footer."""
    document = Document()
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Inches(1))

    normal = document.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    normal.font.size = BODY_SIZE
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15

    style_heading(document, "Heading 1", Pt(14), Pt(16), Pt(6))
    style_heading(document, "Heading 2", Pt(12), Pt(8), Pt(4))

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run(f"Dwayne Aitken | Case Study Solutions: {topic} | Page ")
    add_page_number_field(footer)

    document.core_properties.author = "Dwayne Aitken"
    document.core_properties.title = f"Case Study Solutions: {topic}"
    document.core_properties.comments = ""
    return document


def add_title_block(document, topic, opening):
    """Write the title, the student details and the opening paragraph."""
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(f"Case Study Solutions: {topic}")
    title_run.bold = True
    title_run.font.size = Pt(16)

    add_table(
        document,
        [
            ("Student", "Dwayne Aitken"),
            ("Programme", "Associate Degree in Artificial Intelligence"),
            ("Course", "Probability and Statistics for AI"),
            ("Lesson by", "Prof. Chithra"),
            ("Institution", "Amber Heart Coding Academy"),
            ("Date", "November 4, 2026"),
        ],
        (1.4, 5.1),
        header=False,
        label_column=True,
    )
    document.add_paragraph(opening)


def add_tasks(document, tasks):
    """List each task from the brief, followed directly by its answer."""
    document.add_heading("Tasks", level=2)
    for number, (question, answer) in enumerate(tasks, start=1):
        question_paragraph = document.add_paragraph()
        question_paragraph.paragraph_format.keep_with_next = True
        question_paragraph.paragraph_format.space_after = Pt(2)
        question_paragraph.add_run(f"Task {number}: {question}").bold = True
        document.add_paragraph(answer)


def add_reference(document):
    """End the paper with the APA reference for the lesson notes."""
    document.add_heading("Reference", level=1)
    author, work_title, remainder = REFERENCE_TEXT
    paragraph = document.add_paragraph(author)
    paragraph.paragraph_format.left_indent = Inches(0.5)
    paragraph.paragraph_format.first_line_indent = Inches(-0.5)
    paragraph.add_run(work_title).italic = True
    paragraph.add_run(remainder)


def add_case_opening(document, case):
    """Write the case heading and the scenario sentence from the brief."""
    document.add_heading(case["title"], level=1)
    paragraph = document.add_paragraph(case["brief"])
    paragraph.paragraph_format.keep_with_next = True


def add_given_data_table(document, case):
    """Write the labelled data table supplied in the brief."""
    add_subheading(document, "Given Data")
    rows = [case["headers"]] + list(zip(case["labels"], case["values"]))
    add_table(document, rows, (2.0, 2.2))


def add_subheading(document, text):
    """Write a small section heading inside a case study."""
    document.add_heading(text, level=2)


def add_working(document, rows):
    """Write the required submission items as a label and value table."""
    add_subheading(document, "Working")
    add_table(document, rows, (2.0, 4.5), header=False, label_column=True)


def add_mean_case(document, case):
    """Write one complete mean case study."""
    check_mean(case)
    add_case_opening(document, case)
    add_given_data_table(document, case)
    add_working(
        document,
        [
            ("Total of all values", sum_line(case["values"])),
            ("Number of observations", len(case["values"])),
            ("Formula", MEAN_FORMULA),
            ("Calculation", case["calculation"]),
            ("Final answer", case["final"]),
            ("Interpretation", case["interpretation"]),
        ],
    )
    add_tasks(document, case["tasks"])


def add_median_case(document, case):
    """Write one complete median case study."""
    check_median(case)
    add_case_opening(document, case)
    add_given_data_table(document, case)
    count = len(case["values"])
    rule = MEDIAN_ODD_RULE if count % 2 else MEDIAN_EVEN_RULE
    add_working(
        document,
        [
            ("Ascending order", joined(sorted(case["values"]))),
            ("Number of observations", count),
            ("Middle value(s)", case["middle"]),
            ("Rule", rule),
            ("Median calculation", case["calculation"]),
            ("Final answer", case["final"]),
            ("Interpretation", case["interpretation"]),
        ],
    )
    add_tasks(document, case["tasks"])


def add_mode_case(document, case):
    """Write one complete mode case study."""
    check_mode(case)
    add_case_opening(document, case)
    add_given_data_table(document, case)
    add_subheading(document, "Frequency Table")
    frequencies = frequency_rows(case["values"])
    rows = [(case["value_heading"], "Frequency")]
    rows += frequencies
    rows.append(("Total", sum(frequency for _, frequency in frequencies)))
    add_table(document, rows, (2.0, 2.2))
    add_working(
        document,
        [
            ("Highest frequency", case["highest_frequency"]),
            ("Mode", case["final"]),
            ("Interpretation", case["interpretation"]),
        ],
    )
    add_tasks(document, case["tasks"])


def add_range_case(document, case):
    """Write one complete range case study."""
    check_range(case)
    add_case_opening(document, case)
    add_working(
        document,
        [
            ("Given data", case["data_text"]),
            ("Maximum value", max(case["values"])),
            ("Minimum value", min(case["values"])),
            ("Formula", RANGE_FORMULA),
            ("Calculation", case["calculation"]),
        ],
    )
    add_tasks(document, case["tasks"])


def add_variance_case(document, case):
    """Write one complete population variance case study."""
    check_variance(case)
    add_case_opening(document, case)
    values = case["values"]
    mean = case["mean"]
    add_subheading(document, "Working")
    rows = [(case["value_heading"], "x − Mean", "(x − Mean)²")]
    rows += [(value, value - mean, (value - mean) ** 2) for value in values]
    rows.append(("Total: " + f"{sum(values):,}", "0", f"{case['squared_total']:,}"))
    add_table(document, rows, (2.6, 1.5, 1.5))
    add_table(
        document,
        [
            ("Number of values (N)", len(values)),
            ("Mean", f"{sum(values):,} ÷ {len(values)} = {mean}"),
            ("Formula", VARIANCE_FORMULA),
            (
                "Calculation",
                f"{case['squared_total']:,} ÷ {len(values)} = {case['answer']:,}",
            ),
            ("Final answer", case["final"]),
        ],
        (2.0, 4.5),
        header=False,
        label_column=True,
    )
    add_tasks(document, case["tasks"])


# Each topic: its name, its cases, how to write a case and its opening paragraph.
TOPICS = [
    (
        "Mean",
        MEAN_CASES,
        add_mean_case,
        "This paper solves the five mean case studies. For each case, it presents "
        "the given data, the total of all values, the number of observations, the "
        "formula, the final answer and a one-sentence interpretation, followed by "
        "the answer to each task. The formula follows the lesson notes "
        "(Chithra, n.d.). The AI applications described are possible uses and not "
        "findings from the data.",
    ),
    (
        "Median",
        MEDIAN_CASES,
        add_median_case,
        "This paper solves the five median case studies. For each case, it presents "
        "the given data, the data in ascending order, the number of observations, "
        "the middle value or values, the median calculation and the final answer "
        "with an interpretation, followed by the answer to each task. The method "
        "follows the lesson notes (Chithra, n.d.). The AI applications described "
        "are possible uses and not findings from the data.",
    ),
    (
        "Mode",
        MODE_CASES,
        add_mode_case,
        "This paper solves the five mode case studies. For each case, it presents "
        "the given data, a frequency table, the highest frequency, the mode and a "
        "one-sentence interpretation, followed by the answer to each task. The "
        "method follows the lesson notes (Chithra, n.d.). The AI applications "
        "described are possible uses and not findings from the data.",
    ),
    (
        "Range",
        RANGE_CASES,
        add_range_case,
        "This paper solves the five range case studies. For each case, it presents "
        "the given data, the maximum and minimum values, the formula and the "
        "calculation, followed by the answer to each task. The formula follows the "
        "lesson notes (Chithra, n.d.).",
    ),
    (
        "Variance",
        VARIANCE_CASES,
        add_variance_case,
        "This paper solves the five population variance case studies. For each "
        "case, it presents a table of deviations from the mean, the formula and the "
        "calculation, followed by the answer to each task. Population variance is "
        "used because each brief requests it, and the formula follows the lesson "
        "notes (Chithra, n.d.). Variance is expressed in squared units, so the "
        "standard deviation is also mentioned where it helps the explanation.",
    ),
]


def build_topic(topic, cases, add_case, opening):
    """Build and save the Word document for one statistics topic."""
    document = new_document(topic)
    add_title_block(document, topic, opening)
    for case in cases:
        add_case(document, case)
    add_reference(document)
    path = OUTPUT_FOLDER / f"Case Study {topic} - Dwayne Aitken.docx"
    document.save(path)
    print(f"Saved {path}")


def main():
    """Build all five documents, stopping if any written answer is wrong."""
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    for topic, cases, add_case, opening in TOPICS:
        build_topic(topic, cases, add_case, opening)


if __name__ == "__main__":
    sys.exit(main())
