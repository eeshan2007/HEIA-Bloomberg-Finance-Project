import csv,re,math,collections,statistics,json
from pathlib import Path
from artifact_tool import Workbook,SpreadsheetFile
import openpyxl
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER
ROOT=Path('/mnt/data'); src=openpyxl.load_workbook(ROOT/'HEIA_NA_Equity_Bloomberg_Extract_SYNTHETIC.xlsx',data_only=True)
years=list(range(2017,2025))
def clean(v):
 if v is None:return None
 if isinstance(v,(int,float)):return float(v)
 s=str(v).strip().replace('\u00a0','').replace('\u202f','').replace(' ','').replace(',','').replace('−','-').replace('–','-')
 if s.upper() in ('','—','–','-','N/A','#N/AN/A','NM','N.A.','#N/A'):return None
 neg=s.startswith('(') and s.endswith(')')
 s=s.strip('()');pct=s.endswith('%');s=s.rstrip('%')
 try:return float(s)*(-1 if neg else 1)/(100 if pct else 1)
 except:return None
rows=list(csv.reader(open(ROOT/'HEIA_NA_Equity_RAW_EXPORT.csv',encoding='utf-8-sig',newline='')))
raw=[]; seen=set(); duplicates=[]
for i,r in enumerate(rows,1):
 if not r or len(r)<10:continue
 label=' '.join(r[0].split()); vals=[clean(x) for x in r[2:10]]
 if not label or all(x is None for x in vals):continue
 if i<10 or i>55:continue
 if label.lower()=='net revenue' and label in seen:duplicates.append((i,label));continue
 seen.add(label.lower())
 raw.append((i,label,vals))
lookup={label.lower():(i,v) for i,label,v in raw}
def find(part):
 for i,label,v in raw:
  if label.lower()==part.lower():return v
 raise KeyError(part)
rev=find('Net Revenue');total=find('Total Revenue (incl. Excise)');exc=find('Excise Tax Expense');gp=find('GROSS PROFIT');ebit=find('OPERATING PROFIT (EBIT)');ni=find('NET INCOME ATTRIB. TO SHAREHOLDERS');opex=find('Total Operating Expenses');eps=find('Basic EPS');shares=find('Wtd. Avg Shares - Basic (mn)');cost=find('Cost of Sales');oi=find('Other Income');imp=find('Impairment of PP&E and Goodwill');tax=find('Income Tax Expense');pretax=find('Income Before Taxes');nettotal=find('Net Income (Total)');minor=find('Less: Minority Interest')
# Sign correction: excise displayed as positive but deducted.
raw=[(i,l,[-x if x is not None else None for x in v] if l=='Excise Tax Expense' else v) for i,l,v in raw]
# Inspect export & other for unit outlier
for i,l,v in raw:
 if l=='Export & Other':
  for k,x in enumerate(v):
   if x is not None and x>rev[k]*2:v[k]=x/1e6
# Source tables
seg=src['REV BY SEGMENT'];geo=src['REV BY GEO'];op=src['OPEX DETAIL'];dictionary=src['LOOKUP TABLES']
def vals(s,r):return [clean(s.cell(r,j).value) for j in range(3,11)]
segments=[]
for group,rr in [('Product',range(7,11)),('Brand',range(15,20)),('Channel',range(23,27))]:
 for r in rr:
  name=' '.join(str(seg.cell(r,1).value).split())
  if name=='eB2B / Digital':name='Digital B2B (eB2B)'
  if name=='Export and Other':name='Export & Other'
  for y,v in zip(years,vals(seg,r)):
   if v is not None:segments.append((group,name,y,v))
geo_data=[]
for r in list(range(7,13))+list(range(15,20)):
 name=' '.join(str(geo.cell(r,1).value).split()); vv=vals(geo,r)
 for y,v in zip(years,vv):
  if v is not None:geo_data.append((name,y,v))
regionmap={'Western Europe':[('Europe',1)],'C. & E. Europe':[('Europe',.6),('AMEE',.4)],'Africa & Middle East':[('AMEE',1)],'Americas':[('Americas',1)],'Asia Pacific':[('Asia-Pacific',1)],'Head Office & Elims':[('Head Office & Eliminations',1)],'Europe':[('Europe',1)],'AMEE':[('AMEE',1)],'Asia-Pacific':[('Asia-Pacific',1)],'Head Office & Eliminations':[('Head Office & Eliminations',1)]}
regional=collections.defaultdict(float)
for name,y,v in geo_data:
 for target,weight in regionmap[name]:regional[(target,y)]+=v*weight
operating=[]
for r in range(7,13):
 name=str(op.cell(r,1).value)
 for y,v in zip(years,vals(op,r)):
  if v is not None:operating.append((name,y,v))
volume=vals(op,25);employees=vals(op,24)
# Fix unusual unit by checking gross revenue and channel sums
channels=collections.defaultdict(float)
for group,name,y,v in segments:
 if group=='Channel':channels[y]+=v
checks=[]
def check(name,y,delta,tol=1.1):checks.append((name,y,round(delta,4),abs(delta)<=tol))
for k,y in enumerate(years):
 def v(name):return next((vv[k] for _,l,vv in raw if l.lower()==name.lower()),None)
 check('Channels = net revenue',y,channels[y]-rev[k]);check('Gross less excise = net',y,total[k]-exc[k]-rev[k]);check('Gross profit',y,rev[k]+cost[k]-gp[k]);check('EBIT build',y,gp[k]+oi[k]+opex[k]+(imp[k] or 0)-ebit[k]);check('Pretax build',y,ebit[k]+sum(v(t) or 0 for t in ['Share of Net Profit of Associates & JVs','Interest Income','Interest Expense','Other Net Finance Income/(Expense)'])-pretax[k]);check('Attributable profit',y,nettotal[k]+minor[k]-ni[k]);check('EPS (0.02 tolerance)',y,ni[k]/shares[k]-eps[k],.02)
# write cleaned long csv
fact=[]
for i,label,vv in raw:
 for y,x in zip(years,vv):
  if x is not None:fact.append(['Income statement',label,y,x,'EUR mn' if not ('EPS' in label or 'Per Share' in label or 'Margin' in label or 'Growth' in label or 'Tax Rate' in label or 'Shares' in label) else ('ratio' if '%' in label or 'Growth' in label or 'Tax Rate' in label else 'per share / mn shares'),i])
for group,name,y,v in segments:fact.append([group,name,y,v,'EUR mn',None])
for (name,y),v in sorted(regional.items()):fact.append(['Region restated',name,y,v,'EUR mn',None])
for name,y,v in operating:fact.append(['OPEX detail',name,y,v,'EUR mn',None])
for y,vol,fte in zip(years,volume,employees):
 fact.extend([['Operations','Beer volume (mn hl)',y,vol,'mn hl',25],['Operations','Employees (FTE)',y,fte,'people',24],['Operations','Net revenue per hl',y,rev[y-2017]/vol,'EUR/hl',26],['Operations','Net revenue per employee',y,rev[y-2017]*1e6/fte,'EUR/employee',27]])
with open(ROOT/'HEIA_cleaned_long.csv','w',newline='',encoding='utf-8-sig') as f:
 w=csv.writer(f);w.writerow(['Statement_Group','Line_Item','Fiscal_Year','Value','Unit','Source_Row']);w.writerows(fact)
with open(ROOT/'HEIA_reconciliation.csv','w',newline='') as f:
 w=csv.writer(f);w.writerow(['Check','Year','Difference','Pass']);w.writerows(checks)
# Analysis
metrics=[]
for k,y in enumerate(years):
 metrics.append({'year':y,'revenue':rev[k],'gross':gp[k],'ebit':ebit[k],'ni':ni[k],'opex':-opex[k], 'gm':gp[k]/rev[k], 'om':ebit[k]/rev[k],'nm':ni[k]/rev[k],'eps':eps[k],'volume':volume[k],'eurhl':rev[k]/volume[k], 'fte':employees[k],'eurfte':rev[k]*1e6/employees[k]})
# Workbook
wb=Workbook.create()
for name in ['Executive Summary','Master P&L','Ratio Analysis','YoY Growth','Segment Mix','Restated Regions','Operating Detail','Reconciliation','Data Quality Log','Clean Data','Lookup Dictionary','Chart Data']:wb.worksheets.add(name)
def sh(n):return wb.worksheets.get_item(n)
def put(s,cell,mat):s.get_range(cell).write(mat)
navy='#15253E';blue='#2263A6';pale='#EAF1F9';white='#FFFFFF';grey='#66788B'
def header(s,addr):s.get_range(addr).format={'fill':navy,'font':{'bold':True,'color':white},'row_height':25}
def widths(s,a=31):s.get_range('A:A').format.column_width=a;s.get_range('B:K').format.column_width=15
# Clean Data in A:F, formula lookups by SUMIFS (unique line/year)
cs=sh('Clean Data');put(cs,'A1',[['Statement Group','Line Item','Fiscal Year','Value','Unit','Source Row']]+fact);header(cs,'A1:F1');widths(cs,29);cs.get_range('B:B').format.column_width=43;cs.freeze_panes.freeze_rows(1)
last=len(fact)+1
# master formula SUMIFS guarded by COUNTIFS
ms=sh('Master P&L');put(ms,'A1',[['HEIA NA EQUITY | MASTER P&L | SYNTHETIC DATA']]);put(ms,'A3',[['Line item','Group','Sign']+years]);header(ms,'A3:K3');widths(ms,43);ms.get_range('B:C').format.column_width=19
for ix,(row,label,vv) in enumerate(raw,4):
 put(ms,f'A{ix}',[[label]])
 # dictionary match by normalized code label; exact item dictionary mapped in Python for metadata only
 dictionary_map={'on-trade':('Revenue','Positive'),'off-trade':('Revenue','Positive')}
 group='Revenue' if row<=19 else 'Operating expense' if 27<=row<=33 else 'Financial' if 37<=row<=40 else 'Subtotal'
 sign='Negative' if label in ['Excise Tax Expense','Cost of Sales','Total Operating Expenses','Interest Expense','Income Tax Expense','Less: Minority Interest'] or 27<=row<=33 else 'Positive'
 put(ms,f'B{ix}',[[group,sign]])
 formulas=[[f'=IF(COUNTIFS(\'Clean Data\'!$B$2:$B${last},$A{ix},\'Clean Data\'!$C$2:$C${last},{col}$3)=0,"",SUMIFS(\'Clean Data\'!$D$2:$D${last},\'Clean Data\'!$B$2:$B${last},$A{ix},\'Clean Data\'!$C$2:$C${last},{col}$3))' for col in 'DEFGHIJK']]
 ms.get_range(f'D{ix}:K{ix}').formulas=formulas
 if label.lower() in ['net revenue','gross profit','operating profit (ebit)','net income attrib. to shareholders'] :ms.get_range(f'A{ix}:K{ix}').format.fill=pale
ms.get_range(f'D4:K{3+len(raw)}').set_number_format('#,##0.0;[Red](#,##0.0);–');ms.freeze_panes.freeze_rows(3)
# ratios
rs=sh('Ratio Analysis');put(rs,'A1',[['Profitability & efficiency | 2017–2024']]);put(rs,'A3',[['Metric']+years]);header(rs,'A3:I3');widths(rs,36)
ratio_rows=[('Gross margin',[m['gm'] for m in metrics]),('Operating margin',[m['om'] for m in metrics]),('Net margin',[m['nm'] for m in metrics]),('Gross margin on total revenue',[gp[k]/total[k] for k in range(8)]),('Operating expenses / net revenue',[m['opex']/m['revenue'] for m in metrics]),('Net revenue per hl (EUR)',[m['eurhl'] for m in metrics]),('Net revenue per employee (EUR)',[m['eurfte'] for m in metrics]),('Basic EPS',[m['eps'] for m in metrics])]
for i,(n,values) in enumerate(ratio_rows,4):put(rs,f'A{i}',[[n]+values]);rs.get_range(f'B{i}:I{i}').set_number_format('0.0%' if i<=8 else '#,##0.0')
rs.get_range('B4:I6').conditional_formats.add_color_scale({'minColor':'#FBE6E6','midColor':'#FFFFFF','maxColor':'#D7F2E6'})
ys=sh('YoY Growth');put(ys,'A1',[['Year','Net revenue','Gross profit','EBIT','Attributable net income']]);header(ys,'A1:E1');widths(ys,22)
for k,y in enumerate(years):
 arr=[y]+[(None if k==0 or (seq[k-1] is None or seq[k-1]<=0) else seq[k]/seq[k-1]-1) for seq in [rev,gp,ebit,ni]]
 put(ys,f'A{k+2}',[arr]);ys.get_range(f'B{k+2}:E{k+2}').set_number_format('0.0%')
ss=sh('Segment Mix');put(ss,'A1',[['Dimension','Item','Year','Revenue EUR mn','Share of net revenue']]);header(ss,'A1:E1');widths(ss,25);ss.get_range('B:B').format.column_width=32
for k,(group,name,y,value) in enumerate(segments,2):put(ss,f'A{k}',[[group,name,y,value,value/rev[y-2017]]]);ss.get_range('E2:E250').set_number_format('0.0%')
gs=sh('Restated Regions');put(gs,'A1',[['Comparable reporting region']+years]);header(gs,'A1:I1');widths(gs,35)
for k,name in enumerate(['Europe','Americas','AMEE','Asia-Pacific','Head Office & Eliminations'],2):put(gs,f'A{k}',[[name]+[regional[(name,y)] for y in years]])
put(gs,'A9',[['FY2017–FY2020 C&E Europe allocation: 60% Europe, 40% AMEE; assumption, not actual restated disclosure.']])
os=sh('Operating Detail');put(os,'A1',[['Expense / operating metric']+years]);header(os,'A1:I1');widths(os,37)
for k,r in enumerate([7,8,9,10,11,12,13,24,25],2):put(os,f'A{k}',[[str(op.cell(r,1).value)]+vals(op,r)])
cs2=sh('Reconciliation');put(cs2,'A1',[['Check','Fiscal year','Difference','PASS']]+[[a,b,c,'PASS' if d else 'FAIL'] for a,b,c,d in checks]);header(cs2,'A1:D1');widths(cs2,34)
quality=[['Mixed number formats','Income statement 11–54','NBSP, commas, parentheses, Unicode minus, percentages','Numeric parser; preserve ratio units'],['Repeated Net Revenue','Income statement rows 17, 19','Identical eight-year series','Keep row 17; discard duplicate row 19'],['Repeated Premium','REV BY SEGMENT rows 7, 11','Identical product revenue','Exclude row 11 from product totals'],['Excise sign','Income statement row 16','Source positive; total revenue minus excise = net','Convert to negative in clean data'],['OPEX inclusive impairment','OPEX DETAIL rows 12–13','FY2020/FY2023 total includes impairment','Keep impairment separate in P&L'],['Missing brand history','REV BY SEGMENT rows 15–19','FY2023–FY2024 #N/A N/A','Missing, not zero; latest brand year FY2022'],['Region restatement','REV BY GEO / LOOKUP TABLES','FY2021 reporting change; C&E split 60/40','Assumption-based comparable regions'],['Missing lookup codes','LOOKUP TABLES rows 17,23','Two blank line item codes','Use fallback group; disclose missing'],['Source missing markers','Income statement / OPEX','—, #N/A N/A, NM, n.a.','Distinguish zero/not meaningful/unavailable'],['Possible source unit outlier','Income statement Export & Other','Check channel-to-net revenue reconciliation','Scale only if 1e6 discrepancy confirmed']]
qs=sh('Data Quality Log');put(qs,'A1',[['Issue','Sheet / row','Evidence','Treatment']]+quality);header(qs,'A1:D1');qs.get_range('A:D').format.column_width=37;qs.get_range('C:D').format.column_width=55
ls=sh('Lookup Dictionary');put(ls,'A1',[[dictionary.cell(4,j).value for j in range(1,5)]]+[[dictionary.cell(i,j).value for j in range(1,5)] for i in range(5,37)]);header(ls,'A1:D1');ls.get_range('A:D').format.column_width=31
# chart data
cd=sh('Chart Data');put(cd,'A1',[['Year','Revenue','Operating margin','Gross margin','Net margin','OPEX index','Revenue index','Gross profit index','Volume mn hl','Revenue per hl']]);header(cd,'A1:J1');widths(cd,19)
for k,m in enumerate(metrics,2):put(cd,f'A{k}',[[m['year'],m['revenue'],m['om'],m['gm'],m['nm'],m['opex']/metrics[0]['opex']*100,m['revenue']/metrics[0]['revenue']*100,m['gross']/metrics[0]['gross']*100,m['volume'],m['eurhl']]])
# charts positioned away from tables
for charttype,source,title,pos in [('line','A1:B9','Net revenue (EUR mn)',('L3','S18')),('line','A1:E9','Margins and revenue trend',('L20','S35')),('line','F1:H9','Indexed performance (FY2017=100)',('L37','S52'))]:
 try:
  ch=cd.charts.add(charttype,cd.get_range(source));ch.title_text=title;ch.set_position(*pos)
 except Exception as e:print('chart warning',str(e)[:120])
# executive
es=sh('Executive Summary');put(es,'A1',[['HEIA NA EQUITY | EXECUTIVE OVERVIEW']]);put(es,'A2',[['SYNTHETIC CLASSROOM DATA • FY2017–FY2024 • EUR millions unless stated']]);put(es,'A4',[['Metric','FY2017','FY2024','Change']]);header(es,'A4:D4');widths(es,34)
for k,(name,a,b) in enumerate([('Net revenue',rev[0],rev[-1]),('Gross profit',gp[0],gp[-1]),('Operating profit',ebit[0],ebit[-1]),('Attributable net income',ni[0],ni[-1]),('Gross margin',gp[0]/rev[0],gp[-1]/rev[-1]),('Operating margin',ebit[0]/rev[0],ebit[-1]/rev[-1]),('Basic EPS',eps[0],eps[-1])],5):put(es,f'A{k}',[[name,a,b,b-a]])
put(es,'A14',[['KEY INSIGHTS']]);put(es,'A15',[[f'Net revenue grew {((rev[-1]/rev[0])**(1/7)-1)*100:.1f}% annually, but operating margin moved from {ebit[0]/rev[0]*100:.1f}% to {ebit[-1]/rev[-1]*100:.1f}%.']]);put(es,'A16',[[f'FY2020 EBIT was EUR {ebit[3]:,.0f}mn; FY2024 recovered to EUR {ebit[-1]:,.0f}mn.']]);put(es,'A17',[['Pre-2021 Europe / AMEE comparability relies on an assumed 60/40 regional allocation.']]);es.get_range('A1:D1').format={'fill':navy,'font':{'bold':True,'color':white},'row_height':34};es.get_range('A15:A17').format.column_width=92
SpreadsheetFile.export_xlsx(wb).save(str(ROOT/'HEIA_Completed_Analysis.xlsx'))
# report
styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='Small2',parent=styles['BodyText'],fontSize=9,leading=13,spaceAfter=6));styles.add(ParagraphStyle(name='HeadingBlue',parent=styles['Heading2'],textColor=colors.HexColor(navy),spaceBefore=12));story=[]
def P(s,style='Small2'):story.append(Paragraph(str(s).replace('&','&amp;'),styles[style]))
def H(s):P(s,'HeadingBlue')
def tab(data,widths=None):
 t=Table([[Paragraph(str(c).replace('&','&amp;'),styles['Small2']) for c in row] for row in data],colWidths=widths,repeatRows=1,hAlign='LEFT');t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor(navy)),('TEXTCOLOR',(0,0),(-1,0),colors.white),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F0F5FA')]),('VALIGN',(0,0),(-1,-1),'TOP'),('BOTTOMPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),6)]));story.append(t)
P('HEIA NA Equity | Bloomberg P&amp;L Data Analytics','Title');P('Complete analytical report • FY2017–FY2024 • SYNTHETIC TRAINING DATA','Heading3')
H('Stage 0 — Goal and analytical workflow');P('Decision-maker: CFO / investment analyst evaluating revenue quality, profitability, exposure by channel and region, and operating leverage. Python cleans and reconciles; Excel provides an auditable lookup-based model; Power BI supports filtering and communication. Presentation is analytical risk: confusing total revenue with net revenue understates margins, while treating positive excise as income overstates revenue. The expected discontinuities are pandemic disruption in FY2020 (on-trade and EBIT), inflation / cost pressure after FY2021 (cost of sales and OPEX), and Russia exit / discontinued operations in FY2022–FY2023. Excise is a duty collected on alcoholic products; net revenue excludes it and is the appropriate margin denominator.')
H('Stage 1–2 — Source audit, cleaning and reconciliation');P('Source: 6 workbook sheets (income statement, revenue by segment, geography, OPEX detail, lookup tables, README), plus a CSV mirror. Eight fiscal years. Income-statement data begins at row 10 with headers in rows 7–8 and preamble above. Numeric defects include commas, padding spaces, nonbreaking spaces, parenthetical negatives, Unicode minus, percent strings, and placeholders —, NM, #N/A N/A, n.a. Missing means unavailable, not zero. The duplicated net-revenue row and repeated Premium segment row were deduplicated. The FY2020/FY2023 OPEX-detail total includes impairment, while the income statement separates it. Excise is normalized to a negative deduction. Region names and digital-channel labels are standardized. Missing dictionary codes and regional allocation assumptions are documented in the workbook quality log.')
P(f'Reconciliation results: {sum(x[3] for x in checks)} of {len(checks)} checks pass (tolerance EUR 1.1mn, EPS EUR 0.02). Failed checks must be reviewed in the accompanying reconciliation CSV; differences are not silently forced to zero.')
fail=[x for x in checks if not x[3]]
if fail:tab([['Failed test','Year','Difference']]+[[a,str(b),str(c)] for a,b,c,_ in fail][:15],[240,70,95])
H('Stage 3 — Lookup-driven master model');P('The Excel Master P&amp;L uses SUMIFS with a COUNTIFS guard keyed by exact line item and fiscal year against the long-form Clean Data sheet. This is a two-way lookup: change the row label or year header to retrieve another observation. Blank source values stay blank. The dictionary contains two missing codes; metadata is therefore fallback-tagged rather than represented as a complete code match. Cross-sheet formulas make later refreshes auditable.')
H('Stage 4 — Trends and growth');
cagr=(rev[-1]/rev[0])**(1/7)-1
P(f'Net revenue increased from EUR {rev[0]:,.0f}mn to EUR {rev[-1]:,.0f}mn, a EUR {rev[-1]-rev[0]:,.0f}mn increase and {cagr*100:.2f}% CAGR. Attributable net income moved from EUR {ni[0]:,.0f}mn to EUR {ni[-1]:,.0f}mn; its CAGR is {((ni[-1]/ni[0])**(1/7)-1)*100:.2f}%. Net revenue CAGR is the steadier headline because earnings are disrupted by impairments and below-EBIT items. Growth after FY2020 should be shown with the depressed absolute base; negative prior-year earnings make conventional growth percentages not meaningful.')
tab([['FY','Net revenue','EBIT','Gross margin','Operating margin','EPS']]+[[str(m['year']),f"{m['revenue']:,.0f}",f"{m['ebit']:,.0f}",f"{m['gm']*100:.1f}%",f"{m['om']*100:.1f}%",f"{m['eps']:.2f}"] for m in metrics],[45,86,75,87,103,52])
# volume bridge: delta = (v24-v19)*price19 + v24*(price24-price19)
p19=rev[2]/volume[2];p24=rev[7]/volume[7];vol_effect=(volume[7]-volume[2])*p19;mix_effect=volume[7]*(p24-p19)
P(f'FY2019–FY2024 revenue bridge: change EUR {rev[7]-rev[2]:,.1f}mn = volume effect EUR {vol_effect:,.1f}mn + residual price/mix effect EUR {mix_effect:,.1f}mn, using FY2019 revenue/hl for the volume effect. Price/mix is a residual and includes other factors; it is not a directly observed selling-price series.')
H('Stage 5 — Margins and operating leverage');P(f'FY2017 gross margin {gp[0]/rev[0]*100:.1f}% versus FY2024 {gp[-1]/rev[-1]*100:.1f}%; operating margin {ebit[0]/rev[0]*100:.1f}% versus {ebit[-1]/rev[-1]*100:.1f}%. Gross margin on gross revenue is {gp[-1]/total[-1]*100:.1f}% in FY2024, compared with {gp[-1]/rev[-1]*100:.1f}% on net revenue: a {(gp[-1]/rev[-1]-gp[-1]/total[-1])*100:.1f}pp denominator distortion. Use net revenue for comparability. FY2020 effective tax rate is {(-tax[3]/pretax[3])*100:.1f}% because pretax income was small; avoid interpreting it as a normalized tax burden.')
H('Stage 6 — Segment and geography');
for group in ['Product','Channel']:
 records=[(name,v) for g,name,y,v in segments if g==group and y==2024]
 P(group+' FY2024: '+', '.join(f'{name} EUR {v:,.0f}mn ({v/rev[-1]*100:.1f}%)' for name,v in records)+'.')
P('Brand data are available through FY2022 only; the extract states that from FY2023 disclosure changes to a different presentation. FY2023–FY2024 brand values remain missing, never zero. The restated geography series uses a 60% Europe / 40% AMEE allocation of historical Central &amp; Eastern Europe revenue. Those two regions are estimated before FY2021, not historical reported facts.')
for name in ['Europe','Americas','AMEE','Asia-Pacific','Head Office & Eliminations']:
 P(f'{name}: FY2019 EUR {regional[(name,2019)]:,.1f}mn; FY2024 EUR {regional[(name,2024)]:,.1f}mn; contribution EUR {regional[(name,2024)]-regional[(name,2019)]:+,.1f}mn.')
H('Stage 7–8 — Excel charts and Power BI dashboard');P('The workbook includes a master P&amp;L, ratios, growth table, segment mix, restated regions, operating metrics, reconciliation, quality log, and chart-ready indexed data. Use the companion Power BI implementation guide for model relationships, slicers, DAX measures and visuals. A native .pbix requires Power BI Desktop to create/save and is not generated by this Python workflow.')
H('Stage 9 — Decision-ready insights');
peak=max(metrics,key=lambda m:m['revenue']);gm_hi=max(metrics,key=lambda m:m['gm']);gm_lo=min(metrics,key=lambda m:m['gm']);eps_hi=max(metrics,key=lambda m:m['eps']);eps_lo=min(metrics,key=lambda m:m['eps'])
P(f'Revenue peaked in FY{peak["year"]} at EUR {peak["revenue"]:,.0f}mn. Gross margin peaked in FY{gm_hi["year"]} ({gm_hi["gm"]*100:.1f}%) and bottomed in FY{gm_lo["year"]} ({gm_lo["gm"]*100:.1f}%), a {(gm_hi["gm"]-gm_lo["gm"])*100:.1f}pp spread. EPS peaked in FY{eps_hi["year"]} ({eps_hi["eps"]:.2f}) and bottomed in FY{eps_lo["year"]} ({eps_lo["eps"]:.2f}).')
changes=sorted([(ebit[k]-ebit[k-1],years[k]) for k in range(1,8)])
P(f'Largest EBIT declines: FY{changes[0][1]} (EUR {changes[0][0]:+,.0f}mn) and FY{changes[1][1]} (EUR {changes[1][0]:+,.0f}mn). Interpret FY2020 against pandemic-era on-trade demand and impairment; later weakness also requires examination of cost inflation and below-EBIT financing / discontinued items. These are dataset-consistent interpretations, not causal proof.')
H('One-page CFO memo');P('To: Chief Financial Officer | Subject: Eight-year P&amp;L review — synthetic HEIA NA Equity')
P(f'The eight-year record shows net revenue growth of EUR {rev[-1]-rev[0]:,.0f}mn ({cagr*100:.2f}% CAGR) but less consistent profit conversion. Operating margin moved from {ebit[0]/rev[0]*100:.1f}% in FY2017 to {ebit[-1]/rev[-1]*100:.1f}% in FY2024. FY2020 is a demand and impairment stress test; FY2022–FY2023 must also be read through finance, tax and discontinued operations rather than headline sales alone.')
P(f'Watch next year: (1) operating margin versus FY2024 {ebit[-1]/rev[-1]*100:.1f}% and (2) net revenue per hectolitre versus FY2024 EUR {rev[-1]/volume[-1]:.1f}/hl. Recommendation: prioritize revenue-quality and gross-margin improvement, especially price/mix and route-to-market economics, while tracking OPEX as a share of net revenue. The recommendation is directional: the extract cannot isolate inflation, FX, pricing and portfolio effects causally.')
H('Final reflection');P('The excise sign convention is particularly dangerous: failing to subtract a positive-looking duty would corrupt net revenue and every margin. The unit issue and hidden impairment likewise require reconciliation. Python checks, Excel auditability and Power BI communication together reduce the risk of a persuasive but incorrect presentation.')
doc=SimpleDocTemplate(str(ROOT/'HEIA_Project_Report_and_CFO_Memo.pdf'),pagesize=(595,842),rightMargin=42,leftMargin=42,topMargin=40,bottomMargin=40);doc.build(story)
# BI guide
(ROOT/'PowerBI_Build_Guide_and_DAX.txt').write_text('''HEIA NA Equity — POWER BI DESKTOP IMPLEMENTATION GUIDE (synthetic classroom data)

IMPORT
1. Power BI Desktop > Get Data > Text/CSV > HEIA_cleaned_long.csv. Rename table FactPL.
2. Set Fiscal_Year to whole number, Value to decimal, and the other columns to text.
3. For more robust filtering, create dimensions with DISTINCT: DimYear= DISTINCT(FactPL[Fiscal_Year]); DimItem = DISTINCT(FactPL[Line_Item]); create one-to-many relationships to FactPL. Use Statement_Group and Line_Item as slicers; filter Region restated, Brand, Product, Channel by group.
4. IMPORTANT: The fact table mixes EUR mn, per-share, percentages, volumes and FTE. Always filter Line_Item and Unit in measures. Do not sum all FactPL[Value].

DAX MEASURES
Net Revenue = CALCULATE(SUM(FactPL[Value]), FactPL[Statement_Group] = "Income statement", FactPL[Line_Item] = "Net Revenue")
Gross Profit = CALCULATE(SUM(FactPL[Value]), FactPL[Statement_Group] = "Income statement", FactPL[Line_Item] = "GROSS PROFIT")
EBIT = CALCULATE(SUM(FactPL[Value]), FactPL[Statement_Group] = "Income statement", FactPL[Line_Item] = "OPERATING PROFIT (EBIT)")
Attributable Profit = CALCULATE(SUM(FactPL[Value]), FactPL[Statement_Group] = "Income statement", FactPL[Line_Item] = "NET INCOME ATTRIB. TO SHAREHOLDERS")
Gross Margin = DIVIDE([Gross Profit],[Net Revenue])
Operating Margin = DIVIDE([EBIT],[Net Revenue])
Net Margin = DIVIDE([Attributable Profit],[Net Revenue])
Prior Year Revenue = VAR Y = MAX(FactPL[Fiscal_Year]) RETURN CALCULATE([Net Revenue], REMOVEFILTERS(FactPL[Fiscal_Year]), FactPL[Fiscal_Year] = Y-1)
Revenue Growth = DIVIDE([Net Revenue]-[Prior Year Revenue],[Prior Year Revenue])
OPEX = -CALCULATE(SUM(FactPL[Value]),FactPL[Statement_Group]="Income statement",FactPL[Line_Item]="Total Operating Expenses")
OPEX to Revenue = DIVIDE([OPEX],[Net Revenue])
Beer Volume = CALCULATE(SUM(FactPL[Value]),FactPL[Statement_Group]="Operations",FactPL[Line_Item]="Beer volume (mn hl)")
Revenue per hl = DIVIDE([Net Revenue],[Beer Volume])
EPS = CALCULATE(SUM(FactPL[Value]),FactPL[Statement_Group]="Income statement",FactPL[Line_Item]="Basic EPS")
DPS = CALCULATE(SUM(FactPL[Value]),FactPL[Statement_Group]="Income statement",FactPL[Line_Item]="Dividend Per Share (DPS)")

VISUALS
Top: year-range slicer, latest-year revenue card, growth card, operating-margin card.
Middle: 3-margin line chart; indexed revenue vs OPEX; latest-year total revenue-to-net income waterfall (build with signed individual P&L components, excluding overlapping subtotals); restated-region clustered bars.
Bottom: OPEX composition / net revenue; volume vs EUR/hl combo; EPS vs DPS; three plain-language insight cards.
Use group-specific visuals or dimension tables for region/product/brand/channel: a region filter must not accidentally blank the corporate-wide revenue card. Brand data end in FY2022; label the latest valid year. Pre-2021 Europe and AMEE are estimates using 60/40 allocation.

DELIVERABLE STATUS: This text file is a build guide, NOT a native .pbix file. Open Power BI Desktop and save the finished dashboard as .pbix.
''',encoding='utf-8')
print('FILES CREATED',len(fact),'rows',len(raw),'items',len(checks),'checks',len(fail),'failures',fail[:12]);print('REVENUE',rev,'EBIT',ebit)
