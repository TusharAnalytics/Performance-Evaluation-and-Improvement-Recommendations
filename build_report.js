const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,HeadingLevel,LevelFormat,ImageRun,Header,Footer,PageNumber,PageBreak,BorderStyle}=require('docx');

const FONT="Calibri", NAVY="1F4E79";
const p=(t,o={})=>new Paragraph({spacing:{after:120,line:300},alignment:o.align||AlignmentType.JUSTIFIED,...(o.p||{}),children:[new TextRun({text:t,font:FONT,size:22,...(o.r||{})})]});
const rich=(parts)=>new Paragraph({spacing:{after:120,line:300},alignment:AlignmentType.JUSTIFIED,children:parts.map(([t,b])=>new TextRun({text:t,bold:!!b,font:FONT,size:22}))});
const h1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:320,after:160},children:[new TextRun({text:t,font:FONT,bold:true,size:30,color:NAVY})]});
const h2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:220,after:100},children:[new TextRun({text:t,font:FONT,bold:true,size:25,color:"2E75B6"})]});
const bl=(t,lead)=>new Paragraph({numbering:{reference:"bul",level:0},spacing:{after:80,line:290},children:[...(lead?[new TextRun({text:lead+" ",bold:true,font:FONT,size:22})]:[]),new TextRun({text:t,font:FONT,size:22})]});
const bd={style:BorderStyle.SINGLE,size:4,color:"A6A6A6"}; const borders={top:bd,bottom:bd,left:bd,right:bd};
function table(widths,head,rows,opt={}){
  const total=widths.reduce((a,b)=>a+b,0);
  const cell=(t,i,isH,shade)=>new TableCell({width:{size:widths[i],type:WidthType.DXA},borders,
    shading:{type:ShadingType.CLEAR,fill:isH?NAVY:(shade||"FFFFFF"),color:"auto"},margins:{top:60,bottom:60,left:90,right:90},
    children:[new Paragraph({alignment:AlignmentType.LEFT,children:[new TextRun({text:String(t),font:FONT,size:opt.size||19,bold:isH,color:isH?"FFFFFF":"000000"})]})]});
  return new Table({width:{size:total,type:WidthType.DXA},columnWidths:widths,rows:[
    new TableRow({tableHeader:true,children:head.map((t,i)=>cell(t,i,true))}),
    ...rows.map((r,ri)=>new TableRow({children:r.map((t,i)=>cell(t,i,false,ri%2?"F2F7FB":"FFFFFF"))}))]});
}
const gap=()=>new Paragraph({spacing:{after:120},children:[]});
const cap=t=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:60,after:200},children:[new TextRun({text:t,italics:true,font:FONT,size:19,color:"595959"})]});

const W=9026; // A4 with 1" margins
const children=[];
// Title page
children.push(new Paragraph({spacing:{before:2400,after:200},alignment:AlignmentType.CENTER,children:[new TextRun({text:"SUPPLY CHAIN ANALYTICS INTERNSHIP",font:FONT,size:24,bold:true,color:"7F7F7F"})]}));
children.push(new Paragraph({spacing:{after:200},alignment:AlignmentType.CENTER,children:[new TextRun({text:"Week 6 Report",font:FONT,size:28,color:"7F7F7F"})]}));
children.push(new Paragraph({spacing:{before:400,after:300},alignment:AlignmentType.CENTER,border:{top:{style:BorderStyle.SINGLE,size:12,color:NAVY,space:12},bottom:{style:BorderStyle.SINGLE,size:12,color:NAVY,space:12}},children:[new TextRun({text:"Performance Evaluation and Improvement Recommendations",font:FONT,size:48,bold:true,color:NAVY})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:120},children:[new TextRun({text:"A Data-Driven Assessment of a Hypothetical Distribution Supply Chain",font:FONT,size:26,italics:true})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:1800,after:80},children:[new TextRun({text:"Submitted by: [Your Name]",font:FONT,size:24})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:"Internship Domain: Supply Chain Analytics",font:FONT,size:24})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"October 2026",font:FONT,size:24})]}));
children.push(new Paragraph({children:[new PageBreak()]}));

// Contents
children.push(h1("Contents"));
["1. Introduction","2. Methodology","3. The Hypothetical Supply Chain Scenario","4. Performance Metrics Review","5. Critical Analysis: Strengths, Bottlenecks and Gaps","6. Strategic Recommendations","7. Implementation Roadmap, Risks and Expected Impact","8. Conclusion","References and Assumptions"].forEach(t=>children.push(p(t,{align:AlignmentType.LEFT,p:{spacing:{after:80}}})));
children.push(new Paragraph({children:[new PageBreak()]}));

// 1 Intro
children.push(h1("1. Introduction"));
children.push(p("A supply chain is only as strong as the evidence used to manage it. Over the previous weeks of this internship I collected and cleaned supply chain data (Week 1), identified and analysed key performance indicators (Week 2), mapped processes (Week 3), explored demand forecasting (Week 4) and designed reporting dashboards (Week 5). This final report brings those strands together. It evaluates the performance of a supply chain using a defined set of metrics, interprets what those metrics say about operational efficiency, and proposes practical, justified improvements."));
children.push(p("Because internal company data is not available to an intern, the report is built on a hypothetical but realistic scenario: a mid-sized consumer-goods distributor operating in India. The scenario figures are illustrative and were set to reflect trends commonly reported in publicly available industry material (such as the SCOR framework, logistics cost studies and published inventory and service benchmarks). The goal is not to claim facts about a real firm, but to demonstrate a repeatable method of turning metrics into decisions."));
children.push(h2("1.1 Objectives"));
["Identify and define the KPIs that best describe supply chain health.","Evaluate how each KPI reflects operational efficiency, cost, service and working capital.","Diagnose bottlenecks, gaps and strengths and trace them to root causes.","Recommend prioritised improvements, each justified directly by the analysis.","Present a roadmap, risks and expected financial and service impact."].forEach(t=>children.push(bl(t)));
children.push(h2("1.2 Scope and Limitations"));
children.push(p("The analysis covers the order-to-delivery flow: planning, sourcing, warehousing, transportation and returns. All numbers are hypothetical, rounded and internally consistent. Benchmarks are indicative ranges drawn from general industry practice rather than from a single cited survey, and should be replaced with company-specific or sector-specific values when real data is available."));

// 2 Methodology
children.push(h1("2. Methodology"));
children.push(p("The evaluation followed a five-step, data-driven approach that mirrors common practice in supply chain analytics."));
children.push(table([1100,2600,5326],["Step","Activity","Output"],[
 ["1","KPI selection (SCOR-aligned)","Twelve KPIs across five SCOR attributes: reliability, responsiveness, agility/planning, cost and asset management."],
 ["2","Baseline construction","Hypothetical 12-month averages for each KPI, built from trends in public industry data."],
 ["3","Benchmarking and gap analysis","Gap = benchmark minus current value; status rated Strong, Watch or Critical."],
 ["4","Root-cause analysis","Cause-and-effect (fishbone) reasoning and the 5-Whys technique to link metric gaps to process causes."],
 ["5","Recommendation design and prioritisation","Impact versus effort scoring; targets, roadmap and risk assessment."]
]));
children.push(cap("Table 1: Evaluation methodology"));
children.push(h2("2.1 Rating Rules"));
children.push(bl("within 3% of benchmark (or better).","Strong:"));
children.push(bl("between 3% and 10% behind benchmark.","Watch:"));
children.push(bl("more than 10% behind benchmark, or a gap that directly drives cost or lost sales.","Critical:"));
children.push(h2("2.2 Key Formulas Used"));
[["Perfect Order Rate","(orders delivered on time, in full, undamaged, with correct documents) / total orders"],
 ["OTIF","orders delivered on time and in full / total orders"],
 ["Fill Rate","units shipped immediately from stock / units ordered"],
 ["Inventory Turnover","cost of goods sold / average inventory value"],
 ["Days of Inventory","365 / inventory turnover"],
 ["Forecast Accuracy","1 - MAPE, where MAPE = average of |actual - forecast| / actual"],
 ["Cash-to-Cash Cycle","days of inventory + days sales outstanding - days payables outstanding"],
 ["Logistics Cost %","total logistics cost / net sales"]].forEach(([a,b])=>children.push(bl(b,a+":")));

// 3 Scenario
children.push(h1("3. The Hypothetical Supply Chain Scenario"));
children.push(p("NovaCart Distribution Pvt. Ltd. is an imaginary mid-sized distributor of packaged consumer goods (personal care, household and packaged foods) based in Maharashtra, India. It sources from about 60 manufacturers, operates two regional warehouses and one cross-dock, and serves roughly 4,500 retail outlets and several e-commerce partners."));
children.push(table([3600,5426],["Parameter","Value (illustrative)"],[
 ["Annual net sales","INR 850 crore"],
 ["Cost of goods sold (COGS)","INR 680 crore"],
 ["Active SKUs","About 3,200"],
 ["Suppliers","About 60 (top 10 supply about 65% of volume)"],
 ["Facilities","2 regional warehouses, 1 cross-dock"],
 ["Transport","Mostly outsourced road freight; about 140 delivery routes"],
 ["Planning approach","Monthly spreadsheet-based forecast; limited cross-functional planning"],
 ["Systems","Basic ERP, no integrated transport management or planning system"]
]));
children.push(cap("Table 2: Scenario profile of NovaCart Distribution"));
children.push(p("The company has grown quickly through new retail and online channels, but its processes, systems and planning discipline have not scaled at the same pace. This is a very common situation in growing distributors and makes it a good case for performance evaluation."));

// 4 Metrics
children.push(h1("4. Performance Metrics Review"));
children.push(p("Table 3 presents the baseline values, benchmarks and ratings. Each metric is then discussed in terms of what it measures and why it matters to overall supply chain health."));
children.push(table([2200,1150,1200,1150,3326],["KPI","Current","Benchmark","Status","Significance"],[
 ["Perfect Order Rate","87%","95%","Critical","Overall reliability; one failure in any order dimension breaks the order."],
 ["OTIF","84%","95%","Critical","Retailers penalise misses; links directly to customer retention."],
 ["Fill Rate","93%","97%","Watch","Indicates lost sales and stock availability."],
 ["Forecast Accuracy","68%","80%","Critical","Drives every downstream stocking and capacity decision."],
 ["Supplier On-Time Delivery","82%","93%","Critical","Upstream reliability; feeds stock-outs and expediting."],
 ["Order Cycle Time","6.8 days","4.0 days","Critical","Responsiveness; speed from order to delivery."],
 ["Inventory Turnover","5.2x","8.0x","Critical","Capital efficiency; low turns mean cash locked in stock."],
 ["Days of Inventory","70 days","46 days","Critical","Holding cost and obsolescence exposure."],
 ["Cash-to-Cash Cycle","62 days","40 days","Critical","Working capital and financing cost."],
 ["Logistics Cost (% of sales)","11.5%","8.5%","Critical","Direct margin impact; reflects network and routing efficiency."],
 ["Order Picking Accuracy","98.4%","99.0%","Strong","Warehouse execution quality."],
 ["Return Rate","4.1%","3.0%","Watch","Reverse logistics cost and customer experience."]
],{size:18}));
children.push(cap("Table 3: KPI baseline versus benchmark"));
children.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({type:"png",data:fs.readFileSync("kpi_chart.png"),transformation:{width:560,height:294},altText:{title:"KPI chart",description:"Current vs target vs benchmark for five service KPIs",name:"kpi"}})]}));
children.push(cap("Figure 1: Service and planning KPIs against 12-month targets and benchmarks"));
children.push(h2("4.1 Reliability: Perfect Order, OTIF and Fill Rate"));
children.push(p("Reliability is the most visible dimension to customers. NovaCart's OTIF of 84% means roughly one order in six arrives late or incomplete. Because the perfect order rate (87%) is higher than OTIF, the data suggests that timing and quantity failures are the dominant problem, while documentation and damage errors are smaller contributors. A fill rate of 93% against a 97% benchmark confirms that stock availability is a measurable cause of incomplete deliveries."));
children.push(h2("4.2 Responsiveness: Order Cycle Time"));
children.push(p("An order cycle of 6.8 days compared with a 4-day benchmark shows slow order processing, waiting for stock replenishment and multi-stop routes. Longer cycles force customers to order earlier and in larger lots, which in turn increases forecast noise: a bullwhip effect created in part by the company's own lead times."));
children.push(h2("4.3 Planning: Forecast Accuracy"));
children.push(p("Forecast accuracy of 68% is the single weakest planning indicator. Using spreadsheet-based monthly forecasting without promotion calendars, channel-level signals or exception management means the plan is wrong for roughly one-third of demand. Poor forecasts show up elsewhere in the table as simultaneous stock-outs on fast movers and excess stock on slow movers."));
children.push(h2("4.4 Asset Management: Inventory Turnover, Days of Inventory and Cash-to-Cash"));
children.push(p("Inventory of INR 130.8 crore (COGS 680 / 5.2 turns) supports 70 days of cover. At the 8x benchmark the same sales would need only about INR 85 crore. The cash-to-cash cycle of 62 days includes 70 days of inventory, offset by credit terms, and is 22 days longer than the benchmark. Together these metrics show that the company is financing a large amount of stock that does not translate into better service, since OTIF is still low. High inventory with low availability is the classic symptom of misallocated stock rather than insufficient stock."));
children.push(h2("4.5 Cost: Logistics Cost Percentage"));
children.push(p("Logistics cost is 11.5% of sales (about INR 97.8 crore) against an 8.5% benchmark. The 3 percentage-point gap equals around INR 25.5 crore a year. Contributing factors include partially loaded trucks, many small drops, premium freight used to recover late supplier deliveries, and returns handling."));
children.push(h2("4.6 Execution: Picking Accuracy and Returns"));
children.push(p("Order picking accuracy of 98.4% is close to benchmark, which shows the warehouse team's basic execution is sound. The return rate of 4.1% is moderately high and is partly a consequence of wrong or substituted items and delivery delays rather than product quality alone."));

// 5 Analysis
children.push(h1("5. Critical Analysis: Strengths, Bottlenecks and Gaps"));
children.push(h2("5.1 Strengths"));
children.push(bl("Picking accuracy of 98.4% indicates disciplined warehouse operations and a good foundation for automation or slotting improvements.","Warehouse execution:"));
children.push(bl("A fill rate of 93% shows that the basic product range and stock coverage are broadly adequate; the issue is allocation, not total quantity.","Product availability baseline:"));
children.push(bl("Ownership of two warehouses and a cross-dock gives physical flexibility to rebalance stock and consolidate shipments.","Network footprint:"));
children.push(bl("Top suppliers are concentrated, giving strong leverage for performance contracts and joint planning.","Supplier concentration as leverage:"));
children.push(h2("5.2 Bottlenecks and Gaps"));
children.push(table([2000,3500,3526],["Bottleneck","Evidence from metrics","Likely root cause"],[
 ["Demand planning","Forecast accuracy 68%; high stock yet low OTIF","Spreadsheet monthly forecast; no promotion or channel signals; no S&OP forum"],
 ["Supplier reliability","Supplier on-time 82%; premium freight to recover","No supplier scorecards; single-sourcing on key items; long, variable lead times"],
 ["Inventory allocation","70 days cover but fill rate 93%","Same safety stock policy for all SKUs; no ABC-XYZ segmentation"],
 ["Transportation","Logistics cost 11.5%; cycle time 6.8 days","No route optimisation or load consolidation; fragmented carriers"],
 ["Visibility and systems","Late detection of misses; manual reporting","No control tower, no exception alerts, limited data integration"],
 ["Returns","Return rate 4.1%","Order errors, delays and weak root-cause tracking of returns"]
],{size:18}));
children.push(cap("Table 4: Bottleneck diagnosis linking metrics to causes"));
children.push(h2("5.3 Root-Cause Chain (5-Whys Example)"));
children.push(p("Problem: OTIF is only 84%."));
["Why? Orders are short-shipped or late because stock is not available at the right warehouse at the right time.","Why? Replenishment is triggered late and the wrong quantities are ordered for fast-moving SKUs.","Why? The forecast for those SKUs is inaccurate (68% overall accuracy) and safety stock is set uniformly.","Why? Planning is done monthly in spreadsheets without demand signals from channels or promotions.","Root cause: absence of an integrated planning process (S&OP) and demand-sensing tools."].forEach(t=>children.push(bl(t)));
children.push(h2("5.4 Interdependencies"));
children.push(p("The metrics are not independent. Poor forecasting increases both stock-outs and excess stock; low supplier reliability forces buffer inventory and premium freight; slow cycle time inflates the bullwhip effect, and all of these push logistics cost and working capital higher. This is why the recommendations in the next section are designed as a connected programme, with planning and supplier reliability treated as the foundation for later cost and speed gains."));

// 6 Recs
children.push(h1("6. Strategic Recommendations"));
children.push(p("Seven recommendations are proposed. Each is linked to evidence from the analysis, has a measurable target, and is scored for impact and effort (H = High, M = Medium, L = Low)."));
const recs=[
["R1. Introduce Statistical Forecasting and a Monthly S&OP Process",
 "Replace spreadsheet forecasting with a statistical model (for example, exponential smoothing with seasonality, with machine-learning enhancements for promoted items) at SKU-location level, and create a monthly Sales and Operations Planning meeting with sales, procurement, finance and logistics.",
 "Forecast accuracy is 68% against an 80% benchmark and is the upstream cause of stock imbalance (Section 4.3, Table 4).",
 "Forecast accuracy to 76% in 12 months; OTIF improves by 3 to 4 points.","Impact H, Effort M"],
["R2. Supplier Scorecards, Collaborative Planning and Dual Sourcing",
 "Rate suppliers monthly on on-time delivery, fill and quality; share rolling forecasts with the top 10 suppliers; agree service-level penalties and incentives; qualify a second source for the 15 most critical SKUs.",
 "Supplier on-time delivery is 82% against 93%, and premium freight is used to recover misses (Table 3, Table 4).",
 "Supplier on-time to 90%; premium freight cost reduced by about 30%.","Impact H, Effort M"],
["R3. ABC-XYZ Inventory Segmentation with Differentiated Safety Stock",
 "Classify SKUs by value (ABC) and demand variability (XYZ). Set service targets and safety stock by segment: higher service for AX items, reduced stock or make-to-order for CZ items. Rebalance stock between warehouses weekly.",
 "70 days of inventory coexists with a 93% fill rate, showing misallocation rather than shortage (Section 4.4).",
 "Inventory turnover from 5.2x to 7.0x; about INR 33 crore of working capital released; fill rate to 96%.","Impact H, Effort L to M"],
["R4. Transport Optimisation and Load Consolidation",
 "Implement a lightweight transport management system for route planning and carrier tendering, consolidate multi-drop routes via the cross-dock, move to fixed delivery days per outlet cluster, and shift a share of volume to long-term contracted carriers.",
 "Logistics cost is 11.5% of sales, a gap of about INR 25.5 crore per year, and cycle time is 6.8 days (Section 4.5).",
 "Logistics cost to 9.5% of sales (about INR 17 crore saving); cycle time to 5 days.","Impact H, Effort M to H"],
["R5. Warehouse Slotting and Order Processing Streamlining",
 "Re-slot fast movers near dispatch, introduce wave picking and barcode confirmation at dispatch, and set a same-day order-release rule for orders received before cut-off.",
 "Picking accuracy is strong (98.4%), so the team can adopt the process improvements with low risk; cycle time still lags the benchmark.",
 "Order processing time reduced by about 25%; picking accuracy maintained at 99%.","Impact M, Effort L"],
["R6. Supply Chain Control Tower and Exception-Based Dashboard",
 "Build on the Week 5 dashboard design to create a weekly-then-daily control tower with automated alerts for late purchase orders, stock below safety level, and at-risk deliveries. Assign an owner to each KPI.",
 "Misses are discovered late and reporting is manual; many KPIs are interrelated and need to be monitored together (Section 5.4).",
 "Faster issue detection; reporting effort reduced; supports achievement of all other targets.","Impact M, Effort M"],
["R7. Returns Root-Cause Programme",
 "Record a reason code for every return, review the top reasons monthly, and fix upstream causes such as wrong picks, substitutions and delivery delays. Introduce a quick disposition process to resell returned goods faster.",
 "Return rate is 4.1% against 3.0% and is partly caused by order errors and delays (Section 4.6).",
 "Return rate to 3.3%; lower reverse logistics cost.","Impact M, Effort L"]
];
recs.forEach(r=>{
  children.push(h2(r[0]));
  children.push(rich([["What to do: ",true],[r[1],false]]));
  children.push(rich([["Justification from analysis: ",true],[r[2],false]]));
  children.push(rich([["Expected result: ",true],[r[3]+" ("+r[4]+")",false]]));
});
children.push(h2("6.1 Prioritisation Summary"));
children.push(table([3300,1300,1300,3126],["Recommendation","Impact","Effort","Priority"],[
 ["R3 Inventory segmentation","High","Low-Med","Quick win: start immediately"],
 ["R5 Warehouse slotting and order release","Medium","Low","Quick win: start immediately"],
 ["R7 Returns root-cause programme","Medium","Low","Quick win"],
 ["R1 Forecasting and S&OP","High","Medium","Foundation: months 1 to 6"],
 ["R2 Supplier collaboration","High","Medium","Foundation: months 1 to 9"],
 ["R6 Control tower","Medium","Medium","Enabler: months 3 to 9"],
 ["R4 Transport optimisation","High","Med-High","Longer-term: months 6 to 15"]
]));
children.push(cap("Table 5: Impact and effort prioritisation"));

// 7 roadmap
children.push(h1("7. Implementation Roadmap, Risks and Expected Impact"));
children.push(h2("7.1 Phased Roadmap"));
children.push(table([1900,3300,3826],["Phase","Actions","Milestones"],[
 ["Phase 1: Months 0 to 3","ABC-XYZ segmentation, warehouse re-slotting, returns reason codes, KPI baseline and owners","New safety stock policy live; returns coding at 100%; KPI dashboard v1"],
 ["Phase 2: Months 3 to 9","Statistical forecast pilot on top 300 SKUs, monthly S&OP, supplier scorecards, control tower alerts","Forecast accuracy above 72%; top 10 suppliers on scorecards"],
 ["Phase 3: Months 9 to 18","Transport management system, route redesign, dual sourcing, extension of forecasting to all SKUs","Logistics cost near 9.5%; OTIF above 91%; turnover near 7x"]
],{size:18}));
children.push(cap("Table 6: Three-phase implementation roadmap"));
children.push(h2("7.2 Target KPI Outcomes at 12 to 18 Months"));
children.push(table([2800,1500,1500,3226],["KPI","Current","Target","Main driver"],[
 ["OTIF","84%","91%","R1, R2, R3"],
 ["Perfect Order Rate","87%","92%","R1, R3, R5, R7"],
 ["Forecast Accuracy","68%","76%","R1"],
 ["Inventory Turnover","5.2x","7.0x","R3, R1"],
 ["Cash-to-Cash Cycle","62 days","50 days","R3, R2"],
 ["Logistics Cost %","11.5%","9.5%","R4, R2"],
 ["Order Cycle Time","6.8 days","5.0 days","R5, R4"]
]));
children.push(cap("Table 7: Targets for the improvement programme"));
children.push(h2("7.3 Illustrative Financial Impact"));
children.push(p("Moving inventory turnover from 5.2x to 7.0x lowers average inventory from about INR 130.8 crore to about INR 97.1 crore, releasing about INR 33.6 crore of cash. Reducing logistics cost from 11.5% to 9.5% of sales saves about INR 17 crore a year. Improved OTIF and fill rate would also protect sales that are currently lost to stock-outs, although this benefit is not quantified here because it depends on customer behaviour and margins. These are indicative figures that would need validation with real data and costed implementation budgets."));
children.push(h2("7.4 Risks and Mitigation"));
children.push(table([2700,3000,3326],["Risk","Impact","Mitigation"],[
 ["Poor data quality","Weak forecasts and misleading KPIs","Start with data cleansing; assign data owners; validate KPIs monthly"],
 ["Resistance to change","Low adoption of S&OP and scorecards","Executive sponsor; early quick wins; training and clear incentives"],
 ["Supplier pushback","Limited collaboration","Share forecasts in return for service commitments; phased rollout"],
 ["Technology cost and complexity","Delay or overspend","Pilot first; use modular cloud tools; stage-gate investments"],
 ["Demand volatility","Targets missed","Scenario planning; review safety stock quarterly"]
],{size:18}));
children.push(cap("Table 8: Key risks and mitigation"));

// 8 conclusion
children.push(h1("8. Conclusion"));
children.push(p("The evaluation shows a supply chain that executes well at the warehouse level but underperforms in planning, supplier reliability, inventory allocation and transport efficiency. The central insight is that high inventory (70 days), low service (OTIF 84%) and high logistics cost (11.5% of sales) are different symptoms of the same underlying problems: weak demand planning, unreliable upstream supply and limited visibility."));
children.push(p("The seven recommendations address these causes in a connected sequence. Quick wins in inventory segmentation, warehouse routines and returns analysis build momentum and release cash early, while forecasting, S&OP and supplier collaboration create the foundation for sustained gains in OTIF, cost and agility. If delivered as planned, the programme would lift OTIF to about 91%, release about INR 33 crore of working capital and cut about INR 17 crore of annual logistics cost in the scenario company."));
children.push(p("Personally, this task showed me how KPIs should be read together rather than in isolation, how benchmarking converts raw numbers into priorities, and how every recommendation must be traceable to evidence. These skills build directly on the earlier weeks: data collection, KPI definition, process mapping, forecasting and dashboard design all contribute to a performance evaluation capability."));

children.push(h1("References and Assumptions"));
children.push(h2("Frameworks and Sources of General Practice"));
["ASCM (APICS) SCOR model: framework of Plan, Source, Make, Deliver, Return and Enable and its performance attributes (reliability, responsiveness, agility, costs, asset management).","Chopra, S. and Meindl, P., Supply Chain Management: Strategy, Planning and Operation (textbook): KPI definitions, bullwhip effect and inventory management.","Public industry material on logistics cost as a percentage of sales in India, OTIF practices in retail and consumer goods, and inventory turnover ranges in FMCG distribution.","Ishikawa (fishbone) diagram and 5-Whys technique from standard quality management practice."].forEach(t=>children.push(bl(t)));
children.push(h2("Assumptions"));
["All company data is hypothetical and rounded; benchmarks are indicative.","Financial effects are estimates and exclude implementation cost, tax and inflation.","Industry-standard definitions are used for each KPI as listed in Section 2.2."].forEach(t=>children.push(bl(t)));

const doc=new Document({
 creator:"Intern",title:"Week 6: Performance Evaluation and Improvement Recommendations",
 styles:{default:{document:{run:{font:FONT,size:22}}},
  paragraphStyles:[{id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:30,bold:true,font:FONT,color:NAVY},paragraph:{spacing:{before:320,after:160},outlineLevel:0}},
  {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{size:25,bold:true,font:FONT,color:"2E75B6"},paragraph:{spacing:{before:220,after:100},outlineLevel:1}}]},
 numbering:{config:[{reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:540,hanging:270}}}}]}]},
 sections:[{properties:{page:{margin:{top:1440,right:1440,bottom:1440,left:1440}}},
  headers:{default:new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[new TextRun({text:"Week 6: Performance Evaluation and Improvement Recommendations",font:FONT,size:16,color:"7F7F7F"})]})]})},
  footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"Page ",font:FONT,size:18}),new TextRun({children:[PageNumber.CURRENT],font:FONT,size:18})]})]})},
  children}]
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("/mnt/user-data/outputs/Week6_Performance_Evaluation_and_Recommendations.docx",b);console.log("ok")});
