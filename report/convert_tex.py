# -*- coding: utf-8 -*-
"""开题报告.md -> body.tex(供 ctexart + biblatex/gb7714 主文件 \\input)。
- 仅取 '## 一、' 到 '## 参考文献' 之间的正文
- 内联 [n](含相邻 [n][m])-> \\supercite{key,...}(GB/T7714 上标)
- 代码块(ASCII 闭环)-> \\input{fig_loop.tex};§五 节首插入 \\input{fig_arch.tex}
- 表格 -> tabularx;**粗体**->\\textbf;> 引用 / 列表 / 段落
"""
import re

SRC = "/home/lizhaoxi/.kitty/workspaces/s_9205bdbc/kaiti/report/开题报告.md"
OUT = "/home/lizhaoxi/.kitty/workspaces/s_9205bdbc/kaiti/report/body.tex"

NMAP = {
 1:'feng2025cacm',2:'cao2016offpath',3:'feng2020mixedipid',4:'feng2022pmtud',5:'feng2025pmtudisolation',
 6:'yang2024natwifi',7:'wang2025packetsize',8:'man2020saddns',9:'pan2024tcpspoofing',10:'rfc5961',
 11:'rfc6056',12:'cao2019scent',13:'man2024scad',14:'li2025stateshield',15:'shirokov2018katran',
 16:'fabre2018l4drop',17:'miano2019bpfiptables',18:'graf2020cilium',19:'tetragon2022',20:'falco2018',
 21:'gershuni2019prevail',22:'xu2021k2',23:'sun2024stateembedding',24:'he2023crosscontainer',25:'jin2023epf',
 26:'yang2024kgent',27:'deng2024pentestgpt',28:'fang2024oneday',29:'glazunov2024naptime',30:'zhang2025cybench',
 31:'shao2024nyuctf',32:'zhang2024aprsurvey',33:'seccodegen2025',34:'weng2026hl',35:'romeraparedes2024funsearch',
 36:'novikov2025alphaevolve',37:'liu2024eoh',38:'shinn2023reflexion',39:'chen2024selfdebug',40:'zhang2025dgm',
 41:'wang2023voyager',42:'pathak2017icm',43:'burda2019rnd',44:'lehman2011novelty',45:'mouret2015mapelites',
 46:'schmidhuber2010creativity',47:'kephart2003autonomic',48:'yuan2014selfprotecting',49:'standen2021cyborg',
 50:'castro2025llmacd',51:'gilad2012offpath',52:'rfc2827',
 53:'yang2024sweagent',54:'wang2024openhands',55:'anthropic2025claudecode',
 56:'cursor2024',57:'openai2025codexcli',58:'cognition2024devin',
}

UNI = [("→",r"$\to$"),("↺",r"$\circlearrowleft$"),("×",r"$\times$"),("≥",r"$\geq$"),
       ("≤",r"$\leq$"),("≈",r"$\approx$"),("↑",r"$\uparrow$"),("↓",r"$\downarrow$")]

def unifix(s):
    for a,b in UNI: s=s.replace(a,b)
    return s

def esc(s):
    # 仅转义文本态危险符;$ 已由 unifix 引入,勿动
    for a,b in [("&",r"\&"),("%",r"\%"),("#",r"\#"),("_",r"\_")]:
        s=s.replace(a,b)
    return s

def cites(s):
    def repl(m):
        nums=re.findall(r"\[(\d+)\]", m.group(0))
        keys=[NMAP[int(x)] for x in nums if int(x) in NMAP]
        return r"\supercite{%s}" % ",".join(keys) if keys else m.group(0)
    return re.sub(r"(?:\[\d+\])+", repl, s)

def bold(s):
    return re.sub(r"\*\*(.+?)\*\*", r"\\textbf{\1}", s)

def inline(s):
    return bold(cites(esc(unifix(s))))

lines = open(SRC, encoding="utf-8").read().splitlines()
# 定位正文区间
start = next(i for i,l in enumerate(lines) if l.startswith("## 一、"))
end   = next(i for i,l in enumerate(lines) if l.startswith("## 参考文献"))
lines = lines[start:end]

out=[]
i=0; n=len(lines)
def flush_list(buf, kind):
    if not buf: return
    env = "itemize" if kind=="ul" else "enumerate"
    out.append("\\begin{%s}[leftmargin=2em,itemsep=1pt,topsep=2pt]"%env)
    for it in buf: out.append("  \\item "+inline(it))
    out.append("\\end{%s}"%env)

while i<n:
    s=lines[i].rstrip()
    # 代码块 -> 闭环图
    if s.startswith("```"):
        i+=1
        while i<n and not lines[i].startswith("```"): i+=1
        i+=1
        out.append("\\input{fig_loop.tex}")
        continue
    # 表格
    if s.startswith("|") and i+1<n and re.match(r"^\s*\|[\s:|-]+\|\s*$", lines[i+1]):
        header=[c.strip() for c in s.strip().strip("|").split("|")]
        rows=[header]; i+=2
        while i<n and lines[i].lstrip().startswith("|"):
            rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")]); i+=1
        ncol=max(len(r) for r in rows)
        colspec = ("l " * min(2,ncol) + "X "*(ncol-min(2,ncol))).strip() if ncol>2 else "l "*ncol
        out.append("\\begin{table}[htbp]\\centering\\small")
        out.append("\\begin{tabularx}{\\textwidth}{%s}"%colspec)
        out.append("\\toprule")
        out.append(" & ".join("\\textbf{%s}"%inline(c) for c in header)+" \\\\")
        out.append("\\midrule")
        for r in rows[1:]:
            cells=[inline(r[k]) if k<len(r) else "" for k in range(ncol)]
            out.append(" & ".join(cells)+" \\\\")
        out.append("\\bottomrule\\end{tabularx}\\end{table}")
        continue
    # 标题
    if s.startswith("## "):
        t=s[3:].strip()
        out.append("\\section*{%s}"%inline(t))
        out.append("\\addcontentsline{toc}{section}{%s}"%inline(t))
        if t.startswith("五、"):
            out.append("\\input{fig_arch.tex}")
        i+=1; continue
    if s.startswith("### "):
        t=s[4:].strip()
        out.append("\\subsection*{%s}"%inline(t))
        out.append("\\addcontentsline{toc}{subsection}{%s}"%inline(t))
        i+=1; continue
    # 引用
    if s.startswith(">"):
        out.append("{\\par\\smallskip\\small\\color{black!55}\\kaishu %s\\par\\smallskip}"%inline(s.lstrip("> ").strip()))
        i+=1; continue
    if re.match(r"^\s*---+\s*$", s):
        i+=1; continue
    # 列表
    m=re.match(r"^(\s*)([-*])\s+(.*)$", s)
    if m:
        buf=[m.group(3)]; i+=1
        while i<n:
            mm=re.match(r"^(\s*)([-*])\s+(.*)$", lines[i].rstrip())
            if mm: buf.append(mm.group(3)); i+=1
            else: break
        flush_list(buf,"ul"); continue
    m=re.match(r"^(\s*)(\d+)\.\s+(.*)$", s)
    if m:
        buf=[m.group(3)]; i+=1
        while i<n:
            mm=re.match(r"^(\s*)(\d+)\.\s+(.*)$", lines[i].rstrip())
            if mm: buf.append(mm.group(3)); i+=1
            else: break
        flush_list(buf,"ol"); continue
    if s.strip()=="":
        i+=1; continue
    # 普通段落
    out.append(inline(s.strip()))
    out.append("")
    i+=1

open(OUT,"w",encoding="utf-8").write("\n".join(out))
print("BODY ->", OUT, "lines:", len(out))
print("supercite count:", sum(x.count("\\supercite") for x in out))
