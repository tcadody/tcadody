"""
NAND CTN 터널링 시뮬레이터 — 결과 플롯 생성
Gate | ALO | BOX | O | N | O | Si.channel | TunOx | Si.dot
Erase 동작: V_gate < 0, 전자 Gate → Si.dot 방향 터널링
"""
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.patches import FancyArrowPatch
from matplotlib.ticker import LogLocator, LogFormatter

# ─── 물리 상수 ────────────────────────────────────────────────────
q    = 1.602e-19   # C
hbar = 1.055e-34   # J·s
m0   = 9.109e-31   # kg
eps0 = 8.854e-12   # F/m
kB   = 1.381e-23   # J/K
h    = 6.626e-34   # J·s
eV   = 1.602e-19   # J/eV

# ─── 유효 질량 (단위: m0) ─────────────────────────────────────────
MSTAR = dict(alo=0.20, sio2=0.50, si3n4=0.42, si=0.32)

# ─── 기본 파라미터 ────────────────────────────────────────────────
Vg_default = -18.0   # V (erase 전압)
T_K        = 300     # K
lambda_mfp = 5e-9    # m (Si.channel 평균자유경로)
E_phonon   = 63e-3 * eV  # J (Si 광학 포논 에너지)
t_erase    = 100e-6  # s (erase 시간)
A_dot      = 1e-14   # m² (Si.dot 면적: 100×100 nm²)

# 레이어 정의: (이름, 두께[nm], ε_r, φ_B[eV], mstar)
LAYERS = [
    dict(name='ALO',    d=5e-9,  eps=9.0,  phi=2.8, mstar=MSTAR['alo'],   type='ins'),
    dict(name='BOX',    d=3e-9,  eps=3.9,  phi=3.1, mstar=MSTAR['sio2'],  type='ins'),
    dict(name='O₁',    d=3e-9,  eps=3.9,  phi=3.1, mstar=MSTAR['sio2'],  type='ins'),
    dict(name='N',      d=7e-9,  eps=7.5,  phi=2.0, mstar=MSTAR['si3n4'], type='trap'),
    dict(name='O₂',    d=3e-9,  eps=3.9,  phi=3.1, mstar=MSTAR['sio2'],  type='ins'),
    dict(name='Si.ch',  d=10e-9, eps=11.7, phi=0.0, mstar=MSTAR['si'],    type='semi'),
    dict(name='TunOx',  d=4e-9,  eps=3.9,  phi=3.1, mstar=MSTAR['sio2'],  type='ins'),
]

# ─── 전압 분배 (커패시턴스 분압기) ───────────────────────────────
def voltage_dist(layers, Vtotal):
    sumDE = sum(L['d'] / L['eps'] for L in layers)
    result = []
    for L in layers:
        V_i = Vtotal * (L['d'] / L['eps']) / sumDE
        E_i = V_i / L['d']  # V/m
        result.append({**L, 'V': V_i, 'E_field': E_i})
    return result

# ─── FN 터널링 전류 밀도 [A/m²] ──────────────────────────────────
def fn_current(E_Vm, phi_B_eV, mstar_ratio):
    if E_Vm <= 0: return 1e-100
    phi = phi_B_eV * eV
    meff = mstar_ratio * m0
    A = (q**3) / (8 * np.pi * h * phi) * (m0 / meff)
    B = (4/3) * np.sqrt(2 * meff) * phi**1.5 / (q * hbar)
    return A * E_Vm**2 * np.exp(-B / E_Vm)

# ─── 에너지 손실 모델 (광학 포논 캐스케이드) ─────────────────────
def energy_loss_phonon(E_in_eV, d_ch, lambda_m, E_ph_J, T_K):
    kT = kB * T_K / eV
    E_ph = E_ph_J / eV
    n_max_E = int(E_in_eV / E_ph)
    n_path  = int(d_ch / lambda_m)
    n_col   = min(n_max_E, n_path)
    E_out   = max(E_in_eV - n_col * E_ph, kT)
    return E_out, n_col, E_in_eV - E_out

# ─── Si.channel 내 에너지 프로파일 ───────────────────────────────
def channel_profile(E_in_eV, d_ch, lambda_m, E_ph_J, T_K, N=200):
    kT = kB * T_K / eV
    E_ph = E_ph_J / eV
    xs = np.linspace(0, d_ch, N)
    es = []
    for x in xs:
        n = min(int(x / lambda_m), int(E_in_eV / E_ph))
        E = max(E_in_eV - n * E_ph, kT)
        es.append(E)
    return xs * 1e9, np.array(es)  # nm, eV

# ─── 주어진 전압에서 전류 계산 ───────────────────────────────────
def compute(Vg, T_K=300, lambda_m=5e-9, E_ph_J=63e-3*eV, layers=None):
    if layers is None: layers = LAYERS
    Vtotal = abs(Vg)
    lv = voltage_dist(layers, Vtotal)

    # 차단 스택 (ALO~O₂): 인덱스 0-4
    block = lv[:5]
    semiL = lv[5]
    tunOx = lv[6]

    # 각 차단 레이어의 FN 전류 → 최소값이 병목
    J_block_layers = [fn_current(L['E_field'], L['phi'], L['mstar']) for L in block]
    J_block = min(J_block_layers)

    # Si.channel 입구 운동에너지: 차단 스택 전압 강하
    V_block = sum(L['V'] for L in block)
    E_kin_in = V_block  # eV

    # 에너지 손실
    E_out, n_col, dE = energy_loss_phonon(E_kin_in, semiL['d'], lambda_m, E_ph_J, T_K)

    # TunOx FN (유효 장벽 높이 감소)
    phi_eff = max(tunOx['phi'] - E_out, 0.05)
    J_tox = fn_current(tunOx['E_field'], phi_eff, tunOx['mstar'])

    J_dot = min(J_block, J_tox)
    return dict(lv=lv, J_block=J_block, J_tox=J_tox, J_dot=J_dot,
                E_kin_in=E_kin_in, E_out=E_out, n_col=n_col, dE=dE,
                V_block=V_block, block_J=J_block_layers, semiL=semiL, tunOx=tunOx)

# ─── J-V 스윕 ────────────────────────────────────────────────────
Vg_arr = np.arange(-5, -25.5, -0.5)
JV_block, JV_tox, JV_dot = [], [], []
for Vg in Vg_arr:
    r = compute(Vg)
    JV_block.append(r['J_block'])
    JV_tox.append(r['J_tox'])
    JV_dot.append(r['J_dot'])
JV_block = np.array(JV_block)
JV_tox   = np.array(JV_tox)
JV_dot   = np.array(JV_dot)

# ─── 기본 전압에서 결과 ──────────────────────────────────────────
R = compute(Vg_default)
lv = R['lv']

# ─── λ 스윕 (에너지 손실 민감도) ─────────────────────────────────
lambda_arr = np.linspace(1e-9, 25e-9, 100)
J_vs_lambda = []
E_out_vs_lambda = []
for lam in lambda_arr:
    r2 = compute(Vg_default, lambda_m=lam)
    J_vs_lambda.append(r2['J_dot'])
    E_out_vs_lambda.append(r2['E_out'])
J_vs_lambda   = np.array(J_vs_lambda)
E_out_vs_lambda = np.array(E_out_vs_lambda)

# ─── 에너지 프로파일 ─────────────────────────────────────────────
ch_x, ch_E = channel_profile(R['E_kin_in'], lv[5]['d'], lambda_mfp, E_phonon, T_K)

# ─── 시간 vs 전자 수 ─────────────────────────────────────────────
t_arr = np.linspace(0, t_erase, 300)
N_arr = R['J_dot'] * A_dot * t_arr / q

# ─── 밴드 다이어그램 데이터 ───────────────────────────────────────
def band_diagram_data(lv, Vg):
    """각 레이어의 전도대 에너지를 계산 (Si.dot CB = 0 기준)"""
    xs, CBs = [], []
    x_cum = 0.0
    CB = Vg + lv[0]['phi']   # gate E_F + first barrier
    xs.append(x_cum * 1e9)
    CBs.append(CB)
    for L in lv:
        CB_end = CB - L['V']
        x_cum += L['d']
        xs.append(x_cum * 1e9)
        CBs.append(CB_end)
        CB = CB_end
    return np.array(xs), np.array(CBs)

bd_x, bd_CB = band_diagram_data(lv, Vg_default)

# ─── 레이어 경계 x 위치 (nm) ─────────────────────────────────────
layer_x = [0]
for L in lv:
    layer_x.append(layer_x[-1] + L['d'] * 1e9)
layer_x = np.array(layer_x)
total_nm = layer_x[-1]

# ═══════════════════════════════════════════════════════════════════
# 플롯
# ═══════════════════════════════════════════════════════════════════
plt.style.use('dark_background')
fig = plt.figure(figsize=(16, 12), facecolor='#0f1117')
gs  = gridspec.GridSpec(3, 3, figure=fig, hspace=0.45, wspace=0.38,
                        left=0.06, right=0.97, top=0.93, bottom=0.06)

ax_band   = fig.add_subplot(gs[0, :])    # 상단 전체: 밴드 다이어그램
ax_jv     = fig.add_subplot(gs[1, 0])   # J-V
ax_energy = fig.add_subplot(gs[1, 1])   # Si.channel 에너지
ax_charge = fig.add_subplot(gs[1, 2])   # 전하 vs 시간
ax_field  = fig.add_subplot(gs[2, 0])   # 레이어별 전계/전압
ax_lambda = fig.add_subplot(gs[2, 1])   # λ 민감도
ax_info   = fig.add_subplot(gs[2, 2])   # 결과 요약

COLORS = {
    'ins':  '#3a72c0',
    'trap': '#c0a030',
    'semi': '#30c060',
    'arrow':'#80c8ff',
    'cb':   '#4499ee',
    'ef':   '#50bb88',
    'jdot': '#50dd90',
}

LAYER_BG = {
    'ins':  '#1a2a40',
    'trap': '#2a2010',
    'semi': '#102018',
}

LAYER_NAMES = [L['name'] for L in lv]
LAYER_TYPES = [L['type'] for L in lv]

# ─── (1) 밴드 다이어그램 ─────────────────────────────────────────
ax = ax_band
ax.set_facecolor('#0d1520')

# 레이어 배경 색칠
for i, L in enumerate(lv):
    x0, x1 = layer_x[i], layer_x[i+1]
    c = LAYER_BG.get(L['type'], '#111')
    ax.axvspan(x0, x1, color=c, alpha=0.7, zorder=0)
    ax.axvline(x1, color='#253545', linewidth=0.8, zorder=1)
    mid = (x0 + x1) / 2
    ax.text(mid, -0.6, L['name'], ha='center', va='top', fontsize=8,
            color='#607890', fontfamily='monospace')

# 전도대 (CB) 그리기
for i, L in enumerate(lv):
    x0, x1 = layer_x[i], layer_x[i+1]
    cb0, cb1 = bd_CB[i], bd_CB[i+1]
    c = COLORS[L['type']] if L['type'] != 'semi' else COLORS['semi']
    ax.plot([x0, x1], [cb0, cb1], color=c, linewidth=2.2, zorder=4)
    # Si.channel: 가전자대 (Eg = 1.12 eV)
    if L['type'] == 'semi':
        Eg_Si = 1.12
        ax.plot([x0, x1], [cb0 - Eg_Si, cb1 - Eg_Si],
                color='#208040', linewidth=1.6, linestyle='--', zorder=4)
        ax.fill_between([x0, x1], [cb0 - Eg_Si, cb1 - Eg_Si], [-1.5, -1.5],
                        color='#0a2010', alpha=0.6, zorder=2)

# Gate E_F (수평 점선)
ax.axhline(Vg_default, xmin=0, xmax=layer_x[0]/total_nm + 0.03,
           color='#3070a8', linewidth=1.4, linestyle=':', zorder=5)
ax.text(-1.5, Vg_default, f'E_F(Gate)\n{Vg_default:.0f} eV', ha='right',
        va='center', color='#3070a8', fontsize=7.5)

# Si.dot E_F
ax.axhline(0, xmin=layer_x[-1]/total_nm - 0.02, xmax=1.05,
           color='#30a060', linewidth=1.4, linestyle=':', zorder=5)
ax.text(total_nm + 1.5, 0, 'E_F(Si.dot)\n0 eV', ha='left',
        va='center', color='#30a060', fontsize=7.5)

# 터널링 화살표 (전자 흐름)
y_arrow = Vg_default + 0.5
if y_arrow > -0.5:
    ax.annotate('', xy=(total_nm * 0.95, y_arrow),
                xytext=(total_nm * 0.05, y_arrow),
                arrowprops=dict(arrowstyle='->', color=COLORS['arrow'],
                                lw=1.8, connectionstyle='arc3,rad=0'))
    ax.text(total_nm * 0.5, y_arrow + 0.15, 'e⁻ 터널링 방향 →',
            ha='center', va='bottom', fontsize=8, color=COLORS['arrow'])

# 에너지 손실 표시 (Si.channel)
si_x0, si_x1 = layer_x[5], layer_x[6]
ax.annotate('', xy=(si_x1, R['E_out']),
            xytext=(si_x0, R['E_kin_in']),
            arrowprops=dict(arrowstyle='->', color='#f09030',
                            lw=1.8, connectionstyle='arc3,rad=-0.3'))
ax.text((si_x0 + si_x1)/2 + 1, (R['E_kin_in'] + R['E_out'])/2,
        f'ΔE={R["dE"]:.2f}eV\n({R["n_col"]} phonon)',
        ha='left', va='center', fontsize=7.5, color='#f09030')

ax.set_xlim(-3, total_nm + 3)
ax.set_ylim(-1.0, max(bd_CB) * 1.1 + 0.3)
ax.set_xlabel('위치 (nm)', fontsize=9, color='#8090a8')
ax.set_ylabel('에너지 (eV)', fontsize=9, color='#8090a8')
ax.set_title(f'NAND CTN 에너지 밴드 다이어그램   [V_gate = {Vg_default} V, Erase 동작]',
             fontsize=11, color='#a0c0e0', pad=8)
ax.tick_params(colors='#607080', labelsize=8)
for sp in ax.spines.values(): sp.set_color('#253545')

# ─── (2) J-V 특성 ─────────────────────────────────────────────────
ax = ax_jv
ax.set_facecolor('#0d1520')

mask = JV_block > 1e-40
ax.semilogy(-Vg_arr[mask], JV_block[mask], color='#6090d0', linewidth=1.5,
            linestyle='--', label='J_block (차단스택)')
mask2 = JV_tox > 1e-40
ax.semilogy(-Vg_arr[mask2], JV_tox[mask2], color='#d09030', linewidth=1.5,
            linestyle=':', label='J_TunOx')
mask3 = JV_dot > 1e-40
ax.semilogy(-Vg_arr[mask3], JV_dot[mask3], color=COLORS['jdot'], linewidth=2.2,
            label='J_dot (Si.dot 유입)')
ax.axvline(abs(Vg_default), color='#ff5060', linewidth=1.2, linestyle='-.',
           alpha=0.8, label=f'V_g={Vg_default}V')

ax.set_xlabel('|V_gate| (V)', fontsize=9, color='#8090a8')
ax.set_ylabel('전류 밀도 (A/m²)', fontsize=9, color='#8090a8')
ax.set_title('J-V 터널링 특성', fontsize=10, color='#a0c0e0')
ax.legend(fontsize=7, facecolor='#0d1520', edgecolor='#253545',
          labelcolor='#a0b8c8')
ax.tick_params(colors='#607080', labelsize=8)
ax.set_facecolor('#0d1520')
for sp in ax.spines.values(): sp.set_color('#253545')
ax.grid(True, color='#1e2838', linewidth=0.5, alpha=0.7)

# ─── (3) Si.channel 에너지 프로파일 ──────────────────────────────
ax = ax_energy
ax.set_facecolor('#0d1520')

kT_eV = kB * T_K / eV
ax.plot(ch_x, ch_E, color='#f0a030', linewidth=2.2, label='E_kin(x)')
ax.axhline(kT_eV, color='#406050', linewidth=1.2, linestyle='--',
           label=f'kT={kT_eV*1000:.1f} meV')
ax.axhline(E_phonon / eV, color='#506070', linewidth=1.0,
           linestyle=':', label=f'ℏω_op={E_phonon/eV*1000:.0f} meV')

# 포논 방출 단계 표시
E_ph = E_phonon / eV
n_steps = R['n_col']
for k in range(n_steps + 1):
    x_step = min(k * lambda_mfp * 1e9, ch_x[-1])
    E_step = max(R['E_kin_in'] - k * E_ph, kT_eV)
    if x_step <= ch_x[-1]:
        ax.axvline(x_step, color='#304858', linewidth=0.8, alpha=0.6)

ax.fill_between(ch_x, ch_E, kT_eV, alpha=0.25, color='#f0a030')
ax.scatter([ch_x[0], ch_x[-1]], [R['E_kin_in'], R['E_out']],
           color=['#60ff80', '#ff6040'], s=60, zorder=5)

ax.annotate(f'{R["E_kin_in"]:.2f} eV', xy=(ch_x[0], R['E_kin_in']),
            xytext=(1, R['E_kin_in'] + 0.05), fontsize=7.5, color='#60ff80')
ax.annotate(f'{R["E_out"]:.2f} eV', xy=(ch_x[-1], R['E_out']),
            xytext=(ch_x[-1] - 3, R['E_out'] + 0.05), fontsize=7.5, color='#ff8060')

ax.set_xlabel('Si.channel 내 위치 (nm)', fontsize=9, color='#8090a8')
ax.set_ylabel('전자 운동에너지 (eV)', fontsize=9, color='#8090a8')
ax.set_title('Si.channel 내 에너지 손실\n(광학 포논 캐스케이드)', fontsize=10, color='#a0c0e0')
ax.legend(fontsize=7, facecolor='#0d1520', edgecolor='#253545',
          labelcolor='#a0b8c8')
ax.tick_params(colors='#607080', labelsize=8)
for sp in ax.spines.values(): sp.set_color('#253545')
ax.grid(True, color='#1e2838', linewidth=0.5, alpha=0.7)

# ─── (4) 전하 vs 시간 ─────────────────────────────────────────────
ax = ax_charge
ax.set_facecolor('#0d1520')

ax.plot(t_arr * 1e6, N_arr, color=COLORS['jdot'], linewidth=2.2)
ax.fill_between(t_arr * 1e6, N_arr, alpha=0.2, color=COLORS['jdot'])

# 몇 개 시점 표시
for t_mark in [10, 50, 100]:
    if t_mark <= t_erase * 1e6:
        N_mark = R['J_dot'] * A_dot * t_mark * 1e-6 / q
        ax.scatter([t_mark], [N_mark], color='#ff8060', s=40, zorder=5)
        ax.text(t_mark + 2, N_mark, f'{N_mark:.1f} 개\n@{t_mark}μs',
                fontsize=6.5, color='#d08060', va='center')

ax.set_xlabel('시간 (μs)', fontsize=9, color='#8090a8')
ax.set_ylabel('Si.dot 유입 전자 수', fontsize=9, color='#8090a8')
ax.set_title(f'전자 누적 (Si.dot 100×100 nm²)\nJ={R["J_dot"]:.2e} A/m²', fontsize=10, color='#a0c0e0')
ax.tick_params(colors='#607080', labelsize=8)
for sp in ax.spines.values(): sp.set_color('#253545')
ax.grid(True, color='#1e2838', linewidth=0.5, alpha=0.7)

# ─── (5) 레이어별 전계 / 전압 ─────────────────────────────────────
ax = ax_field
ax.set_facecolor('#0d1520')

names = [L['name'] for L in lv]
V_arr = [L['V'] for L in lv]
E_arr = [L['E_field'] / 1e6 for L in lv]  # MV/m

x_pos = np.arange(len(names))
bar_colors = ['#3a72c0' if L['type'] == 'ins' else
              '#c0a030' if L['type'] == 'trap' else
              '#30c060' for L in lv]

bars = ax.bar(x_pos, V_arr, color=bar_colors, alpha=0.8, edgecolor='#253545')
ax.set_xticks(x_pos)
ax.set_xticklabels(names, rotation=35, ha='right', fontsize=7.5, color='#8090a8')
ax.set_ylabel('전압 강하 (V)', fontsize=9, color='#8090a8')
ax.set_title(f'레이어별 전압/전계 (V_total={abs(Vg_default):.0f}V)', fontsize=10, color='#a0c0e0')

ax2 = ax.twinx()
ax2.plot(x_pos, E_arr, 'o-', color='#f07040', linewidth=1.8, markersize=5,
         label='E (MV/m)')
ax2.set_ylabel('전계 (MV/m)', fontsize=9, color='#f07040')
ax2.tick_params(colors='#f07040', labelsize=8)

ax.tick_params(colors='#607080', labelsize=8)
for sp in ax.spines.values(): sp.set_color('#253545')
for sp in ax2.spines.values(): sp.set_color('#253545')
ax.grid(True, color='#1e2838', linewidth=0.5, alpha=0.7, axis='y')

# ─── (6) λ 민감도 ─────────────────────────────────────────────────
ax = ax_lambda
ax.set_facecolor('#0d1520')

ax2_l = ax.twinx()
mask_l = J_vs_lambda > 1e-40
ax.semilogy(lambda_arr[mask_l] * 1e9, J_vs_lambda[mask_l],
            color=COLORS['jdot'], linewidth=2.2, label='J_dot')
ax2_l.plot(lambda_arr * 1e9, E_out_vs_lambda, color='#f0a030',
           linewidth=1.8, linestyle='--', label='E_out (eV)')
ax.axvline(lambda_mfp * 1e9, color='#ff5060', linewidth=1.2,
           linestyle='-.', label=f'λ={lambda_mfp*1e9:.0f}nm')

ax.set_xlabel('평균자유경로 λ (nm)', fontsize=9, color='#8090a8')
ax.set_ylabel('J_dot (A/m²)', fontsize=9, color='#8090a8')
ax2_l.set_ylabel('E_out (eV)', fontsize=9, color='#f0a030')
ax2_l.tick_params(colors='#f0a030', labelsize=8)
ax.set_title('λ 변화에 따른 전류/에너지\n(에너지 손실 민감도)', fontsize=10, color='#a0c0e0')
ax.tick_params(colors='#607080', labelsize=8)
for sp in ax.spines.values(): sp.set_color('#253545')
for sp in ax2_l.spines.values(): sp.set_color('#253545')
ax.grid(True, color='#1e2838', linewidth=0.5, alpha=0.7)

# 범례 합치기
lines1, labs1 = ax.get_legend_handles_labels()
lines2, labs2 = ax2_l.get_legend_handles_labels()
ax.legend(lines1 + lines2, labs1 + labs2, fontsize=7,
          facecolor='#0d1520', edgecolor='#253545', labelcolor='#a0b8c8')

# ─── (7) 결과 요약 텍스트 ─────────────────────────────────────────
ax = ax_info
ax.set_facecolor('#0d1218')
ax.axis('off')

N_total = R['J_dot'] * A_dot * t_erase / q
I_dot   = R['J_dot'] * A_dot

lines = [
    ('◆ 동작 조건', '#6090c8', True),
    (f'  V_gate = {Vg_default:.1f} V', '#e0e8f0', False),
    (f'  T = {T_K} K, λ = {lambda_mfp*1e9:.0f} nm', '#e0e8f0', False),
    (f'  ℏω_op = {E_phonon/eV*1000:.0f} meV', '#e0e8f0', False),
    ('', '#e0e8f0', False),
    ('◆ 터널링 결과', '#6090c8', True),
    (f'  J_block = {R["J_block"]:.2e} A/m²', '#80c8e0', False),
    (f'  J_TunOx = {R["J_tox"]:.2e} A/m²', '#80c8e0', False),
    (f'  J_dot   = {R["J_dot"]:.2e} A/m²', '#50dd90', False),
    (f'  I_dot   = {I_dot:.2e} A', '#50dd90', False),
    ('', '#e0e8f0', False),
    ('◆ Si.channel 에너지 손실', '#6090c8', True),
    (f'  E_in  = {R["E_kin_in"]:.3f} eV', '#f0a040', False),
    (f'  E_out = {R["E_out"]:.3f} eV', '#f07040', False),
    (f'  ΔE    = {R["dE"]:.3f} eV', '#f07040', False),
    (f'  포논 충돌 = {R["n_col"]} 회', '#f07040', False),
    ('', '#e0e8f0', False),
    ('◆ Erase 결과', '#6090c8', True),
    (f'  면적 = 100×100 nm²', '#c0c8d0', False),
    (f'  시간 = {t_erase*1e6:.0f} μs', '#c0c8d0', False),
    (f'  Si.dot 유입 전자 = {N_total:.2f} 개', '#50dd90', False),
]

y = 0.97
for txt, col, bold in lines:
    weight = 'bold' if bold else 'normal'
    ax.text(0.04, y, txt, transform=ax.transAxes,
            fontsize=8.2, color=col, fontweight=weight,
            fontfamily='monospace', va='top')
    y -= 0.049

# ─── 공통 타이틀 ─────────────────────────────────────────────────
fig.suptitle('NAND CTN 구조 Erase 터널링 시뮬레이션\n'
             'Gate | ALO | BOX | O | N | O | Si.channel | TunOx | Si.dot',
             fontsize=13, color='#b0d0f0', y=0.99, weight='bold')

plt.savefig('/home/user/tcadody/tunneling_result.png', dpi=150,
            bbox_inches='tight', facecolor='#0f1117')
print("저장 완료: tunneling_result.png")
print(f"\n주요 결과 요약 (V_gate={Vg_default}V):")
print(f"  J_block = {R['J_block']:.3e} A/m²")
print(f"  J_TunOx = {R['J_tox']:.3e} A/m²")
print(f"  J_dot   = {R['J_dot']:.3e} A/m²")
print(f"  E_kin 입구 = {R['E_kin_in']:.3f} eV")
print(f"  E_kin 출구 = {R['E_out']:.3f} eV  (손실={R['dE']:.3f} eV, 포논 {R['n_col']}회)")
print(f"  100μs 후 Si.dot 전자수 = {N_total:.2f} 개")
