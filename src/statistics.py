"""WindStats dataclass + compute_all() for MET_MAST_ANALYSIS
IEC 61400-1 Ed.4 formulas applied.
"""
from __future__ import annotations
import math
import numpy as np
import pandas as pd
from dataclasses import dataclass, field, asdict
from scipy.stats import weibull_min, gumbel_r
from typing import List

@dataclass
class WindStats:
    ws100: float = 0.0; ws80: float = 0.0; ws60: float = 0.0
    alpha: float = 0.0; rho: float = 1.225
    weibull_k: float = 0.0; weibull_c: float = 0.0; wpd: float = 0.0
    ti_p90_15: float = 0.0; v50: float = 0.0; v25: float = 0.0
    recovery: float = 0.0
    monthly_ws:  List[float] = field(default_factory=list)
    monthly_wpd: List[float] = field(default_factory=list)
    dominant_dir: str = "N/A"; dominant_dir_pct: float = 0.0
    site_name: str = ""; period: str = ""
    iec_ws_ok: bool = False; iec_wpd_ok: bool = False
    iec_alpha_ok: bool = False; iec_ti_ok: bool = False; iec_v50_ok: bool = False
    def to_dict(self): return asdict(self)
    def iec_assessment(self):
        d = lambda ok: "PASS" if ok else "FAIL"
        return {"ws100":(self.ws100,>6.0",m(self.ws100>=6)),"wpd":(self.wpd,">200",m(self.wpd>=200)),"alpha":(self.alpha,"0.14-0.20",m(0.14<=self.alpha<=0.20)),"ti_p90":(self.ti_p90_15,(<=0.16",m(self.ti_p90_15<=0.16)),"v50":(self.v50,"<=37.5",m(self.v50<=37.5))}

def compute_all(df, site_name="", period=""):
    s = WindStats(site_name=site_name, period=period)
    for c, a in [("WS100","ws100"),("WS80","ws80"),("WS60","ws60")]:
        if c in df.columns: setattr(s, a, round(float(df[c].mean()),4))
    if "Temp_Avg" in df.columns and "BP_Avg" in df.columns:
        s.rho = round(df["BP_Avg"].mean()*100/(287.05*(df["Temp_Avg"].mean()+273.15)),4)
    if s.ws100>0 and s.ws60>0: s.alpha=round(math.log(s.ws100/s.ws60)/math.log(100/60),4)
    if "WS100" in df.columns:
        ws=df["WS100"].dropna().values; ws=ws[ws>0]
        s.wpd=round(0.5*s.rho*float(np.mean(ws**3)),2)
        if len(ws)>100:
            sh,loc,sc=weibull_min.fit(ws,floc=0); s.weibull_k=round(float(sh),4); s.weibull_c=round(float(sc),4)
        mo=df["WS100"].resample("ME").mean()
        if len(mo)==12: s.monthly_ws=[round(float(v),3) for v in mo.values]; s.monthly_wpd=[round(0.5*s.rho*v**3,1) for v in s.monthly_ws]
        s.recovery=round(df["WS100"].notna().mean()*100,1)
    if "WS100_NE_Avg" in df.columns and "WS100_NE_SD" in df.columns:
        df=df.copy(); df["TI"]=df["WS100_NE_SD"]/df["WS100_NE_Avg"].replace(0,np.nan)
        m=(df["WS100_NE_Avg"]>=14)&(df["WS100_NE_Avg"]<16); sub=df.loc[m,"TI"].dropna()
        if len(sub)>5: s.ti_p90_15=round(float(sub.quantile(0.9)),4)
    mc="WS100_NE_Max" if "WS100_NE_Max" in df.columns else "WS100" if "WS100" in df.columns else None
    if mc:
        mm=df[mc].resample("ME").max().dropna().values
        if len(mm)>=6: lg,sc=gumbel_r.fit(mm); s.v50=round(float(gumbel_r.ppf(1-1/50,loc=lg,scale=sc)),2); s.v25=round(float(gumbel_r.ppf(1-1/25,loc=lg,scale=sc)),2)
    wd="WD95_Avg" if "WD95_Avg" in df.columns else None
    if wd:
        wd=df[wd].dropna()%360; idx=((wd+11.25)%360/22.5).astype(int)%16
        ct=np.bincount(idx.values,minlength=16); dm=int(ct.argmax())
        s.dominant_dir=["N","NNE","NE","ENE","E","ESE","SE","SSE","S","SSW","SW","WSW","W","WNW","NW","NNW"][dm]
        s.dominant_dir_pct=round(ct[dm]/ct.sum()*100,2)
    s.iec_ws_ok=s.ws100>=6.0; s.iec_wpd_ok=s.wpd>=200; s.iec_alpha_ok=0.14<=s.alpha<=0.20; s.iec_ti_ok=s.ti_p90_15<=0.16; s.iec_v50_ok=s.v50<=37.5
    return s
