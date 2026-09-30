# What it looks like

These are real screenshots of a report generated with every optional toggle on (`-DemoData -NoGui -Days 30,60,90 -IncludeDeliveryGroups -IncludeApplications -IncludeClientDevices -IncludeTrend`), rendered in a browser, so there is realistic content in every section.

**Header and summary.** The banded green banner is the demo-data watermark mentioned above — it appears only when `-DemoData` was used, so a real report can never be mistaken for a synthetic one. Below it, the Summary table states the environment, when the report was generated, how much session history the site retains, and the site's delivery group and machine counts, followed by the headline figures for the widest requested window and a bar chart comparing unique users and peak concurrency across every requested window.

**A time-series chart, with date labels along the axis.** This is the "Concurrent sessions over time" chart for the 30-day window — the day-by-day, business-hours-shaped oscillation is exactly what you would expect from weekday office usage, and the x-axis is labelled with real calendar dates rather than relative day counts.

**The concurrency distribution table.** This is the report's plain-language reading of the same data as the chart above — it is the section described in more detail below.

**A delivery group breakdown.** Produced by `-IncludeDeliveryGroups`: a horizontal bar chart of unique users per delivery group, followed by a table adding peak concurrency and total sessions for each group.

**A daily activity trend.** Produced by `-IncludeTrend`: unique users per calendar day across the window, again with real dates along the axis.

**The truncation banner.** This is the report's most important element when it appears; how to read it is covered below. The demo dataset shown elsewhere on this page never triggers it, so this screenshot is from a report built with a constructed scenario — a preflight reporting only 45 days of retained session history against a 90-day request — to show what the banner looks like when it fires. 

This is what it looks like on launch, targeting an on-premises Delivery Controller (the default):

Selecting **Citrix Cloud - Commercial** in the Environment dropdown swaps which fields are active: the Delivery Controller and protocol fields grey out, the Citrix Cloud Customer ID / Service Principal ID / Service Principal Secret fields and the "Citrix docs" link become enabled, and the on-premises run-as choice greys out in turn.
