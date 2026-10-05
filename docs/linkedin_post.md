# LinkedIn launch post (draft)

Fill in the [brackets] from your real Nike run before posting. Attach 2–3 deck pages as images (title, football field, benchmark scorecard) and link the live app + GitHub repo in the first comment.

---

I built an equity research engine that turns a ticker into an investment pitch. 📈

**Ticker-to-Thesis** pulls audited financials for a company and its peers straight from SEC filings, then:

🔹 benchmarks it against the industry: growth, margins, ROIC, cash conversion and leverage, with percentile ranks and peer trends
🔹 values it with a DCF (bottom-up WACC from peer betas, implied credit rating), trading comps and bear/base/bull scenarios
🔹 writes the pitch materials: an Excel model with live formulas and a 15-slide deck

First test case: **Nike (NKE) vs Lululemon, Deckers, Crocs, Under Armour, VF Corp and Columbia**.

The model says **[RATING], price target $[PT] vs $[PRICE] ([UPSIDE])**. The interesting part is *why*: [ONE LINE FROM THE BENCHMARK, e.g. "EBIT margin is Xpp below the peer median after falling Ypp in three years, so the whole valuation hinges on how fast margins recover"].

What I learned building it:
• Real filings are messy. Nike doesn't report an operating income line, companies switch revenue tags after accounting changes, and multi-class share counts disappear from cover pages. Handling that, and showing the source filing for every number, was most of the work.
• A price target is only as good as its assumptions, so every input in the Excel model is live and the app has sliders to stress-test them.

Try it: [APP LINK] · Code: [GITHUB LINK]

Built with Python, SEC EDGAR, pandas, openpyxl, python-pptx and Streamlit. Educational project, not investment advice.

#EquityResearch #Valuation #FinancialModeling #InvestmentBanking #Python #DataScience
