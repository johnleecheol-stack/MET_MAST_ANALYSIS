"""
MET_MAST_ANALYSIS - Charts v2 (Korean)
수정사항:
 1. TI 차트: IEC Class A/B를 상수값 → 풍속 함수로 수정 (IEC 61400-1 Ed.4)
 2. Wind Shear: Wind Shear Rose + Wind Shear by Sector 차트 추가
"""
import pandas as pd
import numpy as np
from scipy.stats import weibull_min, gumbel_r
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import math

# 한글 폰트
fm.fontManager.addfont('/home/claude/fonts/NotoSansCJKkr-Regular.ttf')
fm.fontManager.addfont('/home/claude/fonts/NotoSansCJKkr-Bold.ttf')
plt.rcParams['font.family'] = 'Noto Sans CJK KR'
plt.rcParams['axes.unicode_minus'] = False

# ── 데이터 로드 ──
df = pd.read_csv("/mnt/user-data/uploads/권이리_풍황계측데이터_계측기_2025년_.txt",
                 sep='\t', skiprows=417, encoding='utf-8')
col_names = ['Timestamp',
    'WS100_NE_Avg','WS100_NE_SD','WS100_NE_Min','WS100_NE_Max','WS100_NE_Gust',
    'WS100_SW_Avg','WS100_SW_SD','WS100_SW_Min','WS100_SW_Max','WS100_SW_Gust',
    'WS80_NE_Avg','WS80_NE_SD','WS80_NE_Min','WS80_NE_Max','WS80_NE_Gust',
    'WS80_SW_Avg','WS80_SW_SD','WS80_SW_Min','WS80_SW_Max','WS80_SW_Gust',
    'WS60_NE_Avg','WS60_NE_SD','WS60_NE_Min','WS60_NE_Max','WS60_NE_Gust',
    'WS60_SW_Avg','WS60_SW_SD','WS60_SW_Min','WS60_SW_Max','WS60_SW_Gust',
    'Temp_Avg','Temp_SD','Temp_Min','Temp_Max',
    'BP_Avg','BP_SD','BP_Min','BP_Max',
    'RH_Avg','RH_SD','RH_Min','RH_Max',
    'WD95_Avg','WD95_SD','WD95_Gust',
    'WD75_Avg','WD75_SD','WD75_Gust',
    'WD55_Avg','WD55_SD','WD55_Gust','Extra1','Extra2']
df.columns = col_names[:52]
df['Timestamp'] = pd.to_datetime(df['Timestamp'])
df = df.set_index('Timestamp')
df['WS100'] = (df['WS100_NE_Avg'] + df['WS100_SW_Avg']) / 2
df['WS80']  = (df['WS80_NE_Avg']  + df['WS80_SW_Avg'])  / 2
df['WS60']  = (df['WS60_NE_Avg']  + df['WS60_SW_Avg'])  / 2

ws100 = df['WS100'].mean()
ws80  = df['WS80'].mean()
ws60  = df['WS60'].mean()
temp_mean = df['Temp_Avg'].mean()
bp_mean   = df['BP_Avg'].mean()
rho = (bp_mean*100)/(287.05*(temp_mean+273.15))
alpha = math.log(ws100/ws60)/math.log(100/60)

# 색상
DARK_BG='#0D1B2A'; CYAN='#00E5FF'; ORANGE='#FF9800'; YELLOW='#FFD700'
GREEN='#4CAF50'; PINK='#FF4081'; LBLUE='#29B6F6'; WHITE='#FFFFFF'; GRAY='#546E7A'

months_kr = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']
mws  = [7.761,8.832,6.877,6.567,6.282,5.190,5.285,4.285,5.384,5.352,6.743,7.736]
mwpd = [490.5,614.9,326.5,302.3,250.5,151.1,136.4,87.6,137.7,134.7,316.2,458.6]

def dark_axes(ax):
    ax.set_facecolor(DARK_BG)
    for sp in ax.spines.values(): sp.set_color(GRAY)
    ax.tick_params(colors=WHITE, labelsize=11)
    ax.xaxis.label.set_color(WHITE)
    ax.yaxis.label.set_color(WHITE)

# ──────────────────────────────────────────────────────────
# 차트 1: 월별 풍속 (기존과 동일)
# ──────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 5.5), facecolor=DARK_BG)
for ax in axes: dark_axes(ax)
ax = axes[0]
colors_bar = [CYAN if v==max(mws) else (ORANGE if v==min(mws) else LBLUE) for v in mws]
bars = ax.bar(range(12), mws, color=colors_bar, edgecolor='none', width=0.7)
ax.axhline(ws100, color=YELLOW, lw=2, ls='--', label=f'연평균 {ws100:.3f} m/s')
ax.set_xticks(range(12)); ax.set_xticklabels(months_kr, color=WHITE, fontsize=10)
ax.set_ylabel('풍속 (m/s)', color=WHITE)
ax.legend(facecolor='#1a2744', edgecolor=GRAY, labelcolor=WHITE)
ax.set_title('월별 평균 풍속 @ 100m', color=WHITE, fontsize=13, pad=8)
ax.set_ylim(0, 10)
for bar, v in zip(bars, mws):
    ax.text(bar.get_x()+bar.get_width()/2, v+0.1, f'{v:.1f}', ha='center', va='bottom', color=WHITE, fontsize=8.5)
ax2 = axes[1]
h_range = np.linspace(40, 130, 100)
ax2.plot(ws100*(h_range/100)**alpha, h_range, '--', color=YELLOW, lw=1.5, label=f'α={alpha:.3f}')
ax2.plot([ws60,ws80,ws100],[60,80,100],'o',color=CYAN,ms=10,zorder=5)
ax2.plot([ws100*(120/100)**alpha],[120],'D',color=GREEN,ms=9,label='외삽')
ax2.plot([ws100*(150/100)**alpha],[150],'D',color=GREEN,ms=9)
for h,w in [(120,ws100*(120/100)**alpha),(150,ws100*(150/100)**alpha)]:
    ax2.text(w+0.04,h,f'{w:.2f} m/s',color=GREEN,va='center',fontsize=10)
for h,w in zip([60,80,100],[ws60,ws80,ws100]):
    ax2.text(w+0.04,h,f'{w:.3f} m/s',color=CYAN,va='center',fontsize=10)
ax2.set_xlabel('풍속 (m/s)',color=WHITE); ax2.set_ylabel('높이 (m)',color=WHITE)
ax2.set_ylim(40,160)
ax2.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE)
ax2.set_title('수직 바람 프로파일 (Power Law)',color=WHITE,fontsize=13,pad=8)
ax2.grid(axis='x',color=GRAY,alpha=0.3)
plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart1_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart1 done")

# ──────────────────────────────────────────────────────────
# 차트 2: 풍향 Wind Rose (기존과 동일)
# ──────────────────────────────────────────────────────────
dir_freq   = [6.17,6.53,3.94,2.52,2.99,2.26,1.66,3.48,5.03,6.72,6.12,6.37,7.44,24.06,11.01,3.68]
dir_energy = [6.7,5.7,6.0,3.8,2.8,1.3,1.7,2.6,4.5,6.7,5.9,7.6,7.8,20.4,12.6,4.0]
fig, axes = plt.subplots(1, 2, figsize=(14, 6), facecolor=DARK_BG, subplot_kw={'projection':'polar'})
for ax, data, title, color in [
    (axes[0],dir_freq,'빈도 풍배도 (Frequency %)',CYAN),
    (axes[1],dir_energy,'에너지 풍배도 (Energy %)',ORANGE)]:
    ax.set_facecolor(DARK_BG)
    angles = np.deg2rad(np.arange(0,360,22.5))
    bclrs = [color if v==max(data) else (color+'88' if v>np.mean(data) else GRAY+'66') for v in data]
    ax.bar(angles,data,width=np.deg2rad(22.5)*0.85,color=bclrs,edgecolor=DARK_BG,linewidth=0.5)
    ax.set_theta_zero_location('N'); ax.set_theta_direction(-1)
    ax.set_xticks(np.deg2rad([0,45,90,135,180,225,270,315]))
    ax.set_xticklabels(['북(N)','NE','동(E)','SE','남(S)','SW','서(W)','NW'],color=WHITE,fontsize=10)
    ax.yaxis.set_tick_params(labelcolor=GRAY,labelsize=8)
    ax.grid(color=GRAY,alpha=0.3); ax.spines['polar'].set_color(GRAY)
    ax.set_title(title,color=WHITE,fontsize=12,pad=15)
plt.tight_layout()
plt.savefig('/home/claude/chart2_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart2 done")

# ──────────────────────────────────────────────────────────
# 차트 3: 와이블 (기존과 동일)
# ──────────────────────────────────────────────────────────
ws_arr = df['WS100'].dropna().values; ws_arr = ws_arr[ws_arr>0]
sh, lo, sc = weibull_min.fit(ws_arr, floc=0)
v_range = np.linspace(0, 25, 300)
fig, axes = plt.subplots(1, 2, figsize=(14,5.5), facecolor=DARK_BG)
for ax in axes: dark_axes(ax)
ax = axes[0]
hist_vals, bin_edges = np.histogram(ws_arr, bins=np.arange(0,26,1), density=True)
ax.bar((bin_edges[:-1]+bin_edges[1:])/2, hist_vals, width=0.85, color=LBLUE+'99', edgecolor='none', label='실측 분포')
ax.plot(v_range, weibull_min.pdf(v_range,sh,scale=sc), color=CYAN, lw=3, label=f'연간 k={sh:.3f} C={sc:.3f}')
for sn,(k,c,col,ls) in [('봄',(2.226,7.428,GREEN,'--')),('여름',(2.175,5.558,ORANGE,'-.')),
                         ('가을',(2.341,6.572,PINK,':')),('겨울',(2.361,9.137,YELLOW,'-'))]:
    ax.plot(v_range,weibull_min.pdf(v_range,k,scale=c),color=col,lw=2,ls=ls,label=f'{sn} k={k} C={c}')
ax.set_xlabel('풍속 (m/s)',color=WHITE); ax.set_ylabel('확률밀도',color=WHITE)
ax.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE,fontsize=9)
ax.set_title('와이블 분포 (Weibull PDF)',color=WHITE,fontsize=13,pad=8)
ax.set_xlim(0,22); ax.grid(color=GRAY,alpha=0.3)
ax2=axes[1]
cdf_vals=[weibull_min.cdf(v,sh,scale=sc)*100 for v in v_range]
ax2.fill_between(v_range,cdf_vals,alpha=0.3,color=CYAN)
ax2.plot(v_range,cdf_vals,color=CYAN,lw=3,label='연간 CDF')
ax2.axvline(ws100,color=YELLOW,lw=2,ls='--',label=f'Vave={ws100:.3f} m/s')
ax2.set_xlabel('풍속 (m/s)',color=WHITE); ax2.set_ylabel('누적확률 (%)',color=WHITE)
ax2.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE)
ax2.set_title('누적 분포 (CDF)',color=WHITE,fontsize=13,pad=8)
ax2.set_xlim(0,22); ax2.grid(color=GRAY,alpha=0.3)
plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart3_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart3 done")

# ──────────────────────────────────────────────────────────
# 차트 4A: Wind Shear 수직 프로파일 + 계절별 α (기존)
# ──────────────────────────────────────────────────────────
seasonal_alpha={'봄 (MAM)':0.225,'여름 (JJA)':0.250,'가을 (SON)':0.241,'겨울 (DJF)':0.189}
fig, axes = plt.subplots(1,2,figsize=(14,5.5),facecolor=DARK_BG)
for ax in axes: dark_axes(ax)
ax=axes[0]; h2=np.linspace(40,160,200); colors_s=[GREEN,ORANGE,PINK,CYAN]
for (sn,a),col in zip(seasonal_alpha.items(),colors_s):
    ax.plot(ws60*(h2/60)**a,h2,color=col,lw=2.5,label=f'{sn} α={a}')
ax.plot(ws60*(h2/60)**alpha,h2,color=YELLOW,lw=3.5,ls='--',label=f'연평균 α={alpha:.3f}')
for h,w in [(60,ws60),(80,ws80),(100,ws100)]: ax.plot(w,h,'o',color=WHITE,ms=8)
ax.set_xlabel('풍속 (m/s)',color=WHITE); ax.set_ylabel('높이 (m)',color=WHITE)
ax.set_ylim(40,160)
ax.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE,fontsize=9)
ax.set_title('수직 프로파일 계절별 비교',color=WHITE,fontsize=13,pad=8)
ax.grid(color=GRAY,alpha=0.3)
ax2=axes[1]
alpha_vals=list(seasonal_alpha.values())+[alpha]
bars2=ax2.bar(range(5),alpha_vals,color=colors_s+[YELLOW],edgecolor='none',width=0.6)
ax2.axhline(0.20,color='red',lw=2,ls='--',label='IEC 상한 (0.20)')
ax2.axhline(0.14,color=ORANGE,lw=1.5,ls=':',label='IEC 하한 (0.14)')
ax2.set_xticks(range(5)); ax2.set_xticklabels(['봄','여름','가을','겨울','연평균'],color=WHITE,fontsize=10)
ax2.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE)
ax2.set_title('계절별 Wind Shear 멱지수 α',color=WHITE,fontsize=13,pad=8)
for bar,v in zip(bars2,alpha_vals):
    ax2.text(bar.get_x()+bar.get_width()/2,v+0.003,f'{v:.3f}',ha='center',color=WHITE,fontsize=10)
ax2.set_ylim(0,0.35)
plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart4a_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart4a done")

# ──────────────────────────────────────────────────────────
# 차트 4B: Wind Shear Rose + Wind Shear by Sector ★신규★
# ──────────────────────────────────────────────────────────
labels16=['N','NNE','NE','ENE','E','ESE','SE','SSE','S','SSW','SW','WSW','W','WNW','NW','NNW']
labels16_kr=['N(북)','NNE','NE(북동)','ENE','E(동)','ESE','SE(남동)','SSE',
             'S(남)','SSW','SW(남서)','WSW','W(서)','WNW','NW(북서)','NNW']

# 풍향 섹터 분류
wd = df['WD95_Avg'].dropna() % 360
def get_sector_idx(angle):
    return int((angle + 11.25) % 360 / 22.5) % 16
wd_idx = wd.apply(get_sector_idx)

# 각 섹터별 wind shear α 계산
sector_alpha = []
for i in range(16):
    mask = wd_idx == i
    idx = wd_idx[mask].index
    sub = df.loc[df.index.isin(idx)]
    if len(sub) > 20:
        ws100s = sub[['WS100_NE_Avg','WS100_SW_Avg']].mean(axis=1).mean()
        ws60s  = sub[['WS60_NE_Avg','WS60_SW_Avg']].mean(axis=1).mean()
        if ws100s > 0 and ws60s > 0:
            a = math.log(ws100s/ws60s)/math.log(100/60)
        else:
            a = alpha
    else:
        a = alpha
    sector_alpha.append(round(a, 3))

print("Sector alpha:", sector_alpha)
alpha_mean = np.mean(sector_alpha)
alpha_max  = max(sector_alpha)
alpha_min  = min(sector_alpha)

fig, axes = plt.subplots(1, 2, figsize=(14, 6.5), facecolor=DARK_BG,
                          gridspec_kw={'width_ratios':[1,1.3]})

# ── 왼쪽: Wind Shear Rose (극좌표) ──
ax_polar = fig.add_subplot(1, 2, 1, projection='polar', facecolor=DARK_BG)
axes[0].remove(); axes[0] = ax_polar

angles = np.deg2rad(np.arange(0, 360, 22.5))
# 색상: α 값에 따라 그라데이션
norm = plt.Normalize(vmin=0.14, vmax=0.28)
cmap = plt.cm.RdYlGn_r  # 낮을수록 녹색(좋음), 높을수록 빨강(나쁨)
bar_colors_rose = [cmap(norm(a)) for a in sector_alpha]
bars_rose = ax_polar.bar(angles, sector_alpha, width=np.deg2rad(22.5)*0.85,
                          color=bar_colors_rose, edgecolor='#0D1B2A', linewidth=0.8)
ax_polar.set_theta_zero_location('N'); ax_polar.set_theta_direction(-1)
ax_polar.set_xticks(np.deg2rad([0,45,90,135,180,225,270,315]))
ax_polar.set_xticklabels(['북(N)','NE','동(E)','SE','남(S)','SW','서(W)','NW'],color=WHITE,fontsize=9)
ax_polar.yaxis.set_tick_params(labelcolor=GRAY, labelsize=7)
ax_polar.set_ylim(0, 0.35)
ax_polar.grid(color=GRAY, alpha=0.4)
ax_polar.spines['polar'].set_color(GRAY)
ax_polar.set_title('풍향별 Wind Shear Rose\n(α 멱지수)', color=WHITE, fontsize=12, pad=18)

# 컬러바 추가
sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
sm.set_array([])
cbar = fig.colorbar(sm, ax=ax_polar, shrink=0.6, pad=0.1, orientation='vertical')
cbar.ax.yaxis.set_tick_params(color=WHITE, labelsize=8)
plt.setp(cbar.ax.yaxis.get_ticklabels(), color=WHITE)
cbar.set_label('α 값', color=WHITE, fontsize=9)

# ── 오른쪽: Wind Shear by Sector 바 차트 ──
ax2 = axes[1]; dark_axes(ax2)
x = np.arange(16)
bar_colors_sect = [cmap(norm(a)) for a in sector_alpha]
bars_sect = ax2.bar(x, sector_alpha, color=bar_colors_sect, edgecolor='#0D1B2A', width=0.75, linewidth=0.5)
ax2.axhline(alpha_mean, color=YELLOW, lw=2, ls='--', label=f'평균 α={alpha_mean:.3f}')
ax2.axhline(0.20, color='red', lw=2, ls='--', label='IEC 상한 (0.20)', alpha=0.8)
ax2.axhline(0.14, color=GREEN, lw=1.5, ls=':', label='IEC 하한 (0.14)', alpha=0.8)
ax2.set_xticks(x)
ax2.set_xticklabels(labels16_kr, color=WHITE, fontsize=7.5, rotation=45, ha='right')
ax2.set_ylabel('Wind Shear 멱지수 α', color=WHITE)
ax2.set_title('풍향 섹터별 Wind Shear 멱지수', color=WHITE, fontsize=13, pad=8)
ax2.set_ylim(0, 0.58)
ax2.legend(facecolor='#1a2744', edgecolor=GRAY, labelcolor=WHITE, fontsize=9)
ax2.grid(axis='y', color=GRAY, alpha=0.3)
for bar, a_val in zip(bars_sect, sector_alpha):
    ax2.text(bar.get_x()+bar.get_width()/2, a_val+0.003, f'{a_val:.3f}',
             ha='center', va='bottom', color=WHITE, fontsize=7, rotation=90)

# 통계 텍스트
stat_txt = f'Max: {alpha_max:.3f}  /  Min: {alpha_min:.3f}  /  Mean: {alpha_mean:.3f}'
ax2.text(0.5, 0.97, stat_txt, transform=ax2.transAxes, ha='center', va='top',
         color=CYAN, fontsize=9, bbox=dict(facecolor='#1a2744', edgecolor=GRAY, pad=3))

plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart4b_ko.png', dpi=150, bbox_inches='tight', facecolor=DARK_BG)
plt.close(); print("chart4b (shear rose) done")

# ──────────────────────────────────────────────────────────
# 차트 5: TI ★수정★ — IEC Class A/B를 풍속 함수로
# ──────────────────────────────────────────────────────────
df['TI'] = df['WS100_NE_SD']/df['WS100_NE_Avg'].replace(0,np.nan)

fig, axes = plt.subplots(1,2,figsize=(14,5.5),facecolor=DARK_BG)
for ax in axes: dark_axes(ax)

ax=axes[0]
ws_pts = np.arange(3, 20, 2)
ti_p90_curve=[]
for low in ws_pts:
    mask=(df['WS100_NE_Avg']>=low)&(df['WS100_NE_Avg']<low+2)
    ti_p90_curve.append(df.loc[mask,'TI'].quantile(0.90) if mask.sum()>5 else np.nan)

# ★★ IEC 61400-1 Ed.4 풍속 함수 기준선 ★★
ws_line = np.linspace(3, 20, 300)
ti_iec_A = 0.16 * (0.75 + 5.6 / ws_line)   # Class A
ti_iec_B = 0.14 * (0.75 + 5.6 / ws_line)   # Class B
ti_iec_C = 0.12 * (0.75 + 5.6 / ws_line)   # Class C

ax.plot(ws_pts+1, ti_p90_curve, 'o-', color=PINK, lw=3, ms=8, label='TI P90 (측정)', zorder=5)
ax.plot(ws_line, ti_iec_A, color='red',   lw=2.5, ls='--', label='IEC Class A 기준')
ax.plot(ws_line, ti_iec_B, color=ORANGE,  lw=2,   ls=':',  label='IEC Class B 기준')
ax.plot(ws_line, ti_iec_C, color=YELLOW,  lw=1.5, ls='-.',  label='IEC Class C 기준')
ax.set_xlabel('풍속 (m/s)', color=WHITE); ax.set_ylabel('TI P90', color=WHITE)
ax.legend(facecolor='#1a2744', edgecolor=GRAY, labelcolor=WHITE, fontsize=9)
ax.set_title('TI P90 곡선 @ 100m\n(IEC 61400-1 Ed.4 기준)', color=WHITE, fontsize=12, pad=8)
ax.grid(color=GRAY, alpha=0.3); ax.set_ylim(0, 0.6); ax.set_xlim(2, 20)

# 15m/s 포인트 강조
ti_a_15 = 0.16*(0.75+5.6/15)
ti_b_15 = 0.14*(0.75+5.6/15)
ax.annotate(f'@15m/s\nClass A: {ti_a_15:.3f}', xy=(15, ti_a_15), xytext=(13, 0.35),
            color='red', fontsize=8.5,
            arrowprops=dict(arrowstyle='->', color='red', lw=1.2))
ax.annotate(f'@15m/s\n측정 P90: 0.230', xy=(15, 0.230), xytext=(11.5, 0.47),
            color=PINK, fontsize=8.5,
            arrowprops=dict(arrowstyle='->', color=PINK, lw=1.2))

ax2=axes[1]
ws_bin_labels=[f'{i}~{i+2}' for i in range(3,19,2)]
ti_medians,ti_p25,ti_p75=[],[],[]
for low in range(3,19,2):
    mask=(df['WS100_NE_Avg']>=low)&(df['WS100_NE_Avg']<low+2)
    sub=df.loc[mask,'TI'].dropna()
    if len(sub)>5:
        ti_medians.append(sub.median()); ti_p25.append(sub.quantile(0.25)); ti_p75.append(sub.quantile(0.75))
    else:
        ti_medians.append(np.nan); ti_p25.append(np.nan); ti_p75.append(np.nan)
x_pos=range(len(ws_bin_labels))
ax2.bar(x_pos,ti_medians,color=LBLUE+'99',edgecolor='none',width=0.7,label='TI 중앙값')
err_low=[m-l if not np.isnan(m) else 0 for m,l in zip(ti_medians,ti_p25)]
err_hi =[h-m if not np.isnan(m) else 0 for m,h in zip(ti_medians,ti_p75)]
ax2.errorbar(x_pos,ti_medians,yerr=[err_low,err_hi],fmt='none',color=ORANGE,capsize=5,lw=2)

# ★★ 풍속 구간 대표값에서의 IEC 기준선 (계단형) ★★
ws_bin_centers=[low+1 for low in range(3,19,2)]
iec_a_bins=[0.16*(0.75+5.6/ws) for ws in ws_bin_centers]
iec_b_bins=[0.14*(0.75+5.6/ws) for ws in ws_bin_centers]
ax2.step(x_pos, iec_a_bins, where='mid', color='red',   lw=2.5, ls='--', label='IEC Class A')
ax2.step(x_pos, iec_b_bins, where='mid', color=ORANGE,  lw=2,   ls=':',  label='IEC Class B')

ax2.set_xticks(x_pos); ax2.set_xticklabels(ws_bin_labels,color=WHITE,fontsize=9,rotation=45)
ax2.set_ylabel('TI', color=WHITE)
ax2.legend(facecolor='#1a2744', edgecolor=GRAY, labelcolor=WHITE, fontsize=9)
ax2.set_title('풍속 구간별 TI 분포 & IEC 기준', color=WHITE, fontsize=13, pad=8)
ax2.grid(color=GRAY, alpha=0.3)
plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart5_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart5 (TI fixed) done")

# ──────────────────────────────────────────────────────────
# 차트 6: WPD (기존과 동일)
# ──────────────────────────────────────────────────────────
fig,axes=plt.subplots(1,2,figsize=(14,5.5),facecolor=DARK_BG)
for ax in axes: dark_axes(ax)
ax=axes[0]
clrs=[CYAN if v==max(mwpd) else (ORANGE if v==min(mwpd) else LBLUE) for v in mwpd]
bars3=ax.bar(range(12),mwpd,color=clrs,edgecolor='none',width=0.7)
ax.axhline(281.8,color=YELLOW,lw=2,ls='--',label='연평균 281.8 W/m²')
ax.axhline(200,color=GREEN,lw=2,ls=':',label='개발기준 200 W/m²')
ax.set_xticks(range(12)); ax.set_xticklabels(months_kr,color=WHITE,fontsize=10)
ax.set_ylabel('WPD (W/m²)',color=WHITE)
ax.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE)
ax.set_title('월별 풍력 에너지 밀도 (WPD)',color=WHITE,fontsize=13,pad=8)
for bar,v in zip(bars3,mwpd):
    ax.text(bar.get_x()+bar.get_width()/2,v+5,f'{v:.0f}',ha='center',va='bottom',color=WHITE,fontsize=8)
ax2=axes[1]
hwpd=[60,80,100,120,150]
wvals=[0.5*rho*ws60**3,0.5*rho*ws80**3,0.5*rho*ws100**3,
       0.5*rho*(ws100*(120/100)**alpha)**3,0.5*rho*(ws100*(150/100)**alpha)**3]
ax2.barh(range(5),wvals,color=[LBLUE,LBLUE,CYAN,GREEN,YELLOW],edgecolor='none',height=0.6)
ax2.set_yticks(range(5)); ax2.set_yticklabels([f'{h}m' for h in hwpd],color=WHITE,fontsize=11)
ax2.set_xlabel('WPD (W/m²)',color=WHITE)
ax2.axvline(200,color='red',lw=2,ls='--',label='개발기준')
ax2.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE)
ax2.set_title('고도별 풍력 에너지 밀도',color=WHITE,fontsize=13,pad=8)
for i,v in enumerate(wvals): ax2.text(v+3,i,f'{v:.0f}',va='center',color=WHITE,fontsize=10)
plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart6_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart6 done")

# ──────────────────────────────────────────────────────────
# 차트 7: 극한 풍속 (기존과 동일)
# ──────────────────────────────────────────────────────────
monthly_max=df['WS100_NE_Max'].resample('ME').max().dropna().values
lg,sg=gumbel_r.fit(monthly_max)
def vret(T): return gumbel_r.ppf(1-1/T,loc=lg,scale=sg)
fig,axes=plt.subplots(1,2,figsize=(14,5.5),facecolor=DARK_BG)
for ax in axes: dark_axes(ax)
ax=axes[0]
T_r=np.logspace(0.1,2.5,200)
ax.semilogx(T_r,[vret(t) for t in T_r],color=CYAN,lw=3,label='Gumbel 적합')
ax.scatter(range(1,13),sorted(monthly_max),color=ORANGE,s=60,zorder=5,label='월 최대 풍속')
ax.plot(50,vret(50),'*',color=YELLOW,ms=15,zorder=6,label=f'V50={vret(50):.1f} m/s')
ax.axhline(37.5,color='red',lw=2,ls='--',label='IEC Class III (37.5 m/s)')
ax.set_xlabel('재현 기간 (년)',color=WHITE); ax.set_ylabel('극한 풍속 (m/s)',color=WHITE)
ax.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE)
ax.set_title('Gumbel 극한 풍속 곡선',color=WHITE,fontsize=13,pad=8)
ax.set_ylim(0,60); ax.grid(color=GRAY,alpha=0.3)
ax2=axes[1]
hist,edges=np.histogram(monthly_max,bins=10,density=True)
bc=(edges[:-1]+edges[1:])/2
ax2.bar(bc,hist,width=(edges[1]-edges[0])*0.85,color=LBLUE+'99',edgecolor='none',label='월 최대 풍속 분포')
xg=np.linspace(monthly_max.min()-5,monthly_max.max()+10,200)
ax2.plot(xg,gumbel_r.pdf(xg,lg,sg),color=CYAN,lw=3,label='Gumbel PDF')
ax2.axvline(vret(50),color=YELLOW,lw=2.5,ls='--',label=f'V50={vret(50):.1f}')
ax2.axvline(37.5,color='red',lw=2,ls=':',label='IEC Class III (37.5 m/s)')
ax2.set_xlabel('풍속 (m/s)',color=WHITE); ax2.set_ylabel('확률밀도',color=WHITE)
ax2.legend(facecolor='#1a2744',edgecolor=GRAY,labelcolor=WHITE,fontsize=9)
ax2.set_title('Gumbel 분포 적합 결과',color=WHITE,fontsize=13,pad=8)
ax2.grid(color=GRAY,alpha=0.3)
plt.tight_layout(pad=1.5)
plt.savefig('/home/claude/chart7_ko.png',dpi=150,bbox_inches='tight',facecolor=DARK_BG)
plt.close(); print("chart7 done")
print("All KO charts v2 done!")
