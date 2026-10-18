import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
k=["Perfect Order\nRate (%)","OTIF (%)","Supplier On-Time\nDelivery (%)","Forecast\nAccuracy (%)","Fill Rate (%)"]
cur=[87,84,82,68,93]; bm=[95,95,93,80,97]; tg=[92,91,90,76,96]
import numpy as np
x=np.arange(len(k)); w=0.27
fig,ax=plt.subplots(figsize=(8,4.2),dpi=200)
ax.bar(x-w,cur,w,label="Current",color="#C0504D")
ax.bar(x,tg,w,label="12-month target",color="#F2A93B")
ax.bar(x+w,bm,w,label="Industry benchmark",color="#1F4E79")
ax.set_xticks(x); ax.set_xticklabels(k,fontsize=8)
ax.set_ylim(50,100); ax.set_ylabel("Percent")
ax.set_title("Service and Planning KPIs: Current vs Target vs Benchmark",fontsize=10)
ax.legend(fontsize=8,loc="lower right"); ax.spines[["top","right"]].set_visible(False)
for i,(a,b,c) in enumerate(zip(cur,tg,bm)):
    ax.text(i-w,a+0.6,a,ha="center",fontsize=7); ax.text(i,b+0.6,b,ha="center",fontsize=7); ax.text(i+w,c+0.6,c,ha="center",fontsize=7)
plt.tight_layout(); plt.savefig("kpi_chart.png")
