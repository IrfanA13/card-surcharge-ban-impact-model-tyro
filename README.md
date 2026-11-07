# Can Tyro still hit its FY27 guidance after the RBA card changes?

On 1 October 2026 the RBA banned card surcharging and cut the interchange caps. Consumer credit went from 0.80% to 0.30% and debit from 0.20% to 0.16%. Tyro Payments (ASX: TYR) said the change won't hurt its gross profit because it will pass the lower costs on to its merchants. Its guidance for FY27 is $240m to $255m gross profit.

I wanted to test that. I pulled Tyro's numbers out of its reports, matched them with RBA payments data, analysed it in SQL and then built a scenario model in Excel to forecast FY27 to FY29.

## What I found

My Base case comes out at $239.6m, just under the bottom of guidance. Bear is $226.7m and Bull is $255.0m.

To get to the middle of guidance ($247.5m) Tyro's core volume would need to grow about 7.8% if competition takes even 1 bp off its margin. Over the last three years core volume grew 2.4%, 1.7% and 4.4%.

Tyro has been growing profit by charging more, not by processing more. Its share of card spending in Australia fell from 5.3% to 4.1% between H1 FY23 and H2 FY26, while its payments margin went up from 40.7 to 46.2 bps.

Tyro charges merchants about 97 bps. The RBA market average is about 62 bps. Now that merchants can't surcharge, and fees have to be published from April 2027, small businesses are going to notice that gap. Every 1 bp of margin is worth about $4.4m of gross profit a year.

"Neutral" doesn't mean nothing changes. If Tyro passes on the full interchange saving (about 9.6 bps) its revenue drops around 5% in FY27 even though gross profit stays the same.

## How I thought about it

### Why this project

I picked Tyro because it's a listed Australian payments company with public numbers, its business is mostly one thing (card payments) so it's easy to model, and the RBA reform had just happened. Tyro also made a claim I could test, that the reform is neutral for its gross profit. Most portfolio projects analyse the past. I liked that this one makes a forecast I can check when the real results come out.

### Breaking the business down

The first thing I needed to understand was how Tyro makes money. For every $100 a customer taps at a Tyro merchant, the merchant pays Tyro about 97c in card fees, plus around 7c in other fees like terminal rental. Tyro then pays about 56c to the banks and card schemes, which is mostly interchange. What's left, about 46c, is Tyro's margin.

So payments gross profit is basically volume x margin. That became the backbone of the model. The reform cuts the 56c cost line, and the question is who keeps that saving. If Tyro passes it all to merchants the margin stays the same and only revenue drops. That's why Tyro calls it neutral.

### What I thought Tyro wasn't saying

When I looked at the history in SQL two things stood out. Tyro's volume was flat while the whole card market grew about 29%, so it has been losing share. And its margin kept going up, so the profit growth came from pricing. When I compared Tyro's fee to the RBA market average it was about 35 bps higher.

That made me think the real risk isn't the interchange cut, it's what happens to Tyro's pricing. Before the reform merchants could pass card fees on to customers as a surcharge, so they didn't feel them much. Now they have to absorb them, and from April 2027 providers have to publish their fees. I added a "margin lost to competition" input to the model to test this. Even 1 to 3 bps makes a real difference.

### Why half years and scenarios

Tyro's business is seasonal because H1 (Jul to Dec) has Christmas in it. So I forecast each half from the same half last year instead of from the half before.

I didn't want to give one number because nobody knows exactly how merchants will react. So I built Bear, Base and Bull scenarios plus a sensitivity grid that shows FY27 profit for every mix of volume growth and margin squeeze. That way the model shows which assumptions actually matter instead of pretending to be precise.

### Things I questioned along the way

I wondered whether cafes would just raise their menu prices instead of surcharging. That would actually help Tyro a little because it earns a % of every sale, so higher prices mean more dollars going through its machines. When I sized it, it was only about $1m to $1.5m a year, small compared to the $4.4m from each 1 bp of margin. So I kept it as an optional input rather than a main driver.

I also asked myself how I'd know if the model is any good. I backtested it by pretending it was a year earlier and only using the data I would have had then. It came within about 2% both times. But when I looked closer, in FY26 my volume guess was too low and my margin guess was too high and they partly cancelled out. So I don't just trust the total, I watch volume and margin separately.

### Why SQL and Excel

I used SQL for the history because it's how company data is usually stored and it makes joining Tyro's numbers with the RBA data simple. I used Excel for the forecast because that's what finance teams use for scenario models, and anyone can change an input and see what happens.

## What I used

Python to clean the data and build the tables, SQL (SQLite) for the analysis and Excel for the forecast model.

The data is 6 Tyro investor presentations and 3 annual reports covering H1 FY23 to FY26, plus three RBA tables: C1.1 credit cards, C2.1 debit cards and C3 merchant fees.

## How it works

1. `src/build_dataset.py` builds 5 clean tables. Tyro only publishes PDFs so I typed their numbers in with the page each one came from. The RBA numbers are read straight from the Excel files in `data/raw`. It also runs checks to make sure my numbers line up with what Tyro reported.

2. `sql/week1_queries.sql` has 7 questions I answered in SQL: market share, margin trend, credit vs debit mix, Tyro's price vs the market, which industries are growing, what 1 bp of margin is worth, and how much profit comes from payments. `src/run_sql.py` runs them all and prints the results.

3. `src/build_model.py` creates `Tyro_RBA_Reform_Model.xlsx`. It forecasts each half year from FY27 to FY29 for Bear, Base and Bull, with a Custom column to play with. There's also a sensitivity grid, the growth needed to hit guidance, and a budget vs actual tab for when Tyro reports H1 FY27.

## Is it accurate?

The data checks out. When I work out the margins myself they match what Tyro reported for every year and half year, including the second halves I had to calculate.

The model has 750 formulas and no errors, and I got the same Base result when I calculated it separately in Python.

Backtesting gave errors of +1.9% for FY25 and -0.9% for FY26 on total gross profit.

The real test is Tyro's H1 FY27 result in late February 2027. I'll put the actual numbers in the Budget_vs_Actual tab and see how far off I was and whether it was volume or margin.

## What's next

I'm building a Power BI dashboard from the same tables to show the main results on one page: the scenarios against guidance, market share, Tyro's price vs the market and the sensitivity grid.

After Tyro reports H1 FY27 I'll update the model with the actual numbers and the first RBA data from after the reform.

## Run it yourself

```
pip install pandas openpyxl
python3 src/build_dataset.py
python3 src/run_sql.py
python3 src/build_model.py
```

Then open `Tyro_RBA_Reform_Model.xlsx` in Excel.

## Folders

`data/raw` has the RBA files. `data` has my clean tables as csv files and the SQLite database `tyro_project.db`. `sql` has the queries and `src` has the Python scripts.

## Things to keep in mind

Tyro doesn't publish how much of its volume is credit vs debit so I used the national credit share (39.6%). The interchange rates before the reform are my own estimates and they're highlighted yellow in the Inputs tab.

The market share uses all card spending in Australia, including places Tyro doesn't really compete like big supermarkets and online. So I'd read it as a trend rather than an exact share.

The RBA changed what it includes in the merchant fee data from the December 2024 quarter, so the market average isn't perfectly like for like over time.

There's no data from after the reform yet. The RBA card data goes to August 2026 and the fee data to June 2026.

## Sources

Tyro Payments annual reports FY24 to FY26 and investor presentations H1 FY24 to FY26 (ASX: TYR), available on Tyro's investor centre. Reserve Bank of Australia statistical tables C1.1, C2.1 and C3, used under the RBA's Creative Commons licence. RBA Review of Merchant Card Payment Costs and Surcharging, Conclusions Paper, March 2026.

This is my own analysis using public information and it's not investment advice.

Irfan Akram, Sydney
