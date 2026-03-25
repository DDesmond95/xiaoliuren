# Xiao Liu Ren Explorer

A **fully interactive Xiao Liu Ren divination explorer and research tool** built as a **static website**.

The project combines:

- traditional Chinese divination concepts
- a large generated dataset
- educational explanations
- statistical exploration tools
- a beginner-friendly UI

The entire site runs **purely in the browser** using:

- HTML
- CSS
- JavaScript

No backend server is required.

The site can be deployed directly to **GitHub Pages**.

---

# Demo

When deployed:

```
https://your-username.github.io/xiao-liu-ren-explorer/
```

---

# What is Xiao Liu Ren

Xiao Liu Ren (小六壬) is a **simplified folk divination method** derived from the classical Chinese divination system **Da Liu Ren (大六壬)**.

It is commonly used for **quick situational divination**, such as:

- travel timing
- meeting outcomes
- short-term decisions
- interpersonal situations
- daily fortune checks

The system uses:

```
Lunar month
Lunar day
Earthly branch hour
```

to generate one of **six symbolic results**.

---

# The Six Results

| Result | Meaning                      |
| ------ | ---------------------------- |
| 大安   | Stability, calm, favorable   |
| 留连   | Delay, repetition, lingering |
| 速喜   | Fast progress, good news     |
| 赤口   | Conflict, argument, tension  |
| 小吉   | Moderate good fortune        |
| 空亡   | Emptiness, uncertainty       |

These six results form a **cyclic sequence**.

---

# Calculation Formula

The commonly used formula is:

```
(lunar_month + lunar_day + branch_hour − 3) mod 6
```

Then map the result index to the six results:

```
0 → 大安
1 → 留连
2 → 速喜
3 → 赤口
4 → 小吉
5 → 空亡
```

Example:

```
Lunar month = 3
Lunar day = 23
Branch hour = 8

(3 + 23 + 8 − 3) mod 6
= 31 mod 6
= 1

Result → 留连
```

---

# Website Features

The website is designed for **three audiences**:

```
Beginners
Practitioners
Researchers
```

---

# Beginner Features

The site explains the system clearly.

### Introduction

Explains:

- what Xiao Liu Ren is
- historical background
- how it works

### Glossary

Defines:

- lunar calendar
- earthly branches
- branch hour
- yin-yang
- five elements

### Learning Sections

Includes:

```
step-by-step calculation explanation
cycle diagram
branch hour clock
example interpretations
```

---

# Daily View

The **Today page** automatically displays:

```
current date
current time
current branch hour
lunar date
```

It shows the **24 hourly results** for today.

Features include:

```
hourly cards
timeline view
dominant result summary
highlighted current hour
copy summary button
refresh button
```

---

# Date Lookup

Users can select any date.

The site displays:

```
all hourly results
branch hour
lunar month
lunar day
meaning
advice
```

Additional features:

```
filter by result
filter by branch
timeline or card view
copy summary
```

---

# Manual Xiao Liu Ren Calculator

Users can calculate results directly.

Inputs:

```
lunar month
lunar day
branch hour
```

The result panel shows:

```
final result
classification
meaning
advice
simple explanation
calculation formula
step-by-step breakdown
```

---

# Gregorian → Lunar Conversion

Users can input a normal calendar date.

Inputs:

```
Gregorian date
time
```

The site automatically converts:

```
Gregorian date → lunar date
clock time → branch hour
```

Then calculates the result.

---

# Current Moment Divination

Users can click:

```
Ask Now
```

The site calculates using:

```
current system time
current lunar date
current branch hour
```

This produces a **real-time divination result**.

---

# Compare Two Times

Users can compare two different times.

Example:

```
meeting today vs tomorrow
```

The comparison panel shows:

```
date/time
lunar date
branch hour
result
```

---

# Dataset Explorer

The site loads a **pre-generated dataset**.

Each row includes:

```
datetime
hour
lunar_month
lunar_day
branch
result
meaning
advice
```

Users can:

```
search
filter by result
filter by branch
filter by date range
sort results
paginate rows
view row details
```

---

# Data Export

Filtered results can be downloaded as:

```
JSON
CSV
```

---

# Statistics Dashboard

The statistics page analyzes the dataset.

It shows:

```
result frequency
percentages
hourly distribution
branch distribution
result ranking
favorable vs challenging totals
```

Charts include:

```
percentage bars
hour distribution chart
branch distribution chart
ranking list
```

---

# Visual Learning Tools

The website includes visual aids:

### Six Result Cycle

Displays:

```
大安
留连
速喜
赤口
小吉
空亡
```

The active result is highlighted.

---

### Earthly Branch Clock

Shows the 12 branches:

```
子 丑 寅 卯 辰 巳 午 未 申 酉 戌 亥
```

The current branch is highlighted.

---

# Calculation Transparency

The site shows full calculation steps.

Example breakdown:

```
Lunar month = 3
Lunar day = 23
Branch hour = 8

Compute:
(3 + 23 + 8 − 3)

Raw total = 31
Modulo 6 = 1

Result → 留连
```

This ensures users understand **exactly how results are produced**.

---

# Psychology and Scientific Context

The website also explains psychological aspects of divination.

Topics include:

```
confirmation bias
pattern recognition
uncertainty reduction
symbolic interpretation
decision framing
```

The site emphasizes:

```
cultural tradition
educational value
non-scientific predictive claims
```

---

# Dataset Generation

The dataset is generated using a Python script.

It computes results for every hour between two dates.

Example generation range:

```
2024-05-01 → 2026-01-01
```

Each hour is calculated using:

```
solar date → lunar date
hour → branch
Xiao Liu Ren formula
```

The result is saved as:

```
xiao_liuren_data.json
```

---

# Project Structure

```
repository
│
├─ README.md
├─ generate_dataset.py
│
├─ docs
│  ├─ index.html
│  ├─ styles.css
│  ├─ app.js
│  ├─ xiao_liuren_data.json
│
└─ LICENSE
```

---

# Running Locally

Simply open:

```
docs/index.html
```

in a browser.

No server is required.

---

# Deploying to GitHub Pages

1. Push repository to GitHub

2. Open repository settings

3. Go to:

```
Settings → Pages
```

4. Set:

```
Source → Deploy from branch
Branch → main
Folder → /docs
```

5. Save.

The site will be available at:

```
https://username.github.io/repository-name/
```

---

# Browser Compatibility

The site works in modern browsers:

```
Chrome
Firefox
Edge
Safari
```

The lunar conversion uses:

```
Intl.DateTimeFormat Chinese calendar
```

If unavailable, the site falls back to dataset lookup.

---

# Accessibility

The site includes:

```
keyboard navigation
skip links
ARIA attributes
focus states
high contrast colors
reduced motion support
```

---

# License

MIT License

You may freely:

```
use
modify
distribute
```

with attribution.

---

# References

Chinese metaphysics and divination:

```
Da Liu Ren (大六壬)
Qi Men Dun Jia (奇门遁甲)
Tai Yi (太乙)
```

General background:

```
Chinese calendrical system
Earthly Branches
Traditional divination methods
```

Psychology:

```
confirmation bias
pattern recognition
uncertainty reduction
```

---

# Disclaimer

This project is intended for:

```
education
cultural exploration
software experimentation
```

It should not be treated as:

```
scientific prediction
guaranteed future outcomes
professional advice
```
