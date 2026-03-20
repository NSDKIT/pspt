"""
ドライラン結果の可視化スクリプト
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np

# 日本語フォント設定
plt.rcParams['font.family'] = 'DejaVu Sans'

# ============================================================
# データ定義（ドライラン結果）
# ============================================================
k1_values = list(range(1, 16))          # 1〜15 サイクル
bus_list  = list(range(11, 48))         # バス 11〜47
n_k1      = len(k1_values)
n_bus     = len(bus_list)
cases_per_k1 = n_bus * 37              # バス × バス組み合わせ数

# 総ケース数の内訳
labels = [f'{k}cyc' for k in k1_values]
case_counts = [cases_per_k1] * n_k1

# ============================================================
# Figure 1: 総ケース数の内訳（棒グラフ）
# ============================================================
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
fig.suptitle('CPAT Dry Run - Analysis Case Summary', fontsize=14, fontweight='bold')

ax1 = axes[0]
colors = plt.cm.Blues(np.linspace(0.4, 0.9, n_k1))
bars = ax1.bar(labels, case_counts, color=colors, edgecolor='white', linewidth=0.5)
ax1.set_xlabel('k1 (Short Circuit Duration [cycles])', fontsize=11)
ax1.set_ylabel('Number of Cases', fontsize=11)
ax1.set_title('Cases per k1 Value\n(Total: 20,535 cases)', fontsize=11)
ax1.tick_params(axis='x', rotation=45)
ax1.set_ylim(0, max(case_counts) * 1.2)
ax1.axhline(y=cases_per_k1, color='red', linestyle='--', alpha=0.5, label=f'{cases_per_k1} cases/k1')
ax1.legend(fontsize=9)
# 合計を表示
ax1.text(0.98, 0.95, f'Total: {sum(case_counts):,} cases',
         transform=ax1.transAxes, ha='right', va='top',
         fontsize=11, color='darkblue',
         bbox=dict(boxstyle='round,pad=0.3', facecolor='lightblue', alpha=0.5))

# ============================================================
# Figure 2: .pop パラメータ書き換え結果（テストケース）
# ============================================================
ax2 = axes[1]

test_k1   = [1,  5,  10, 15]
test_bus  = [11, 20, 30, 11]
test_yco  = [0.25, 0.25, 0.50, 0.25]
pgo_val   = 0.038889  # 全ケース共通（Pgo/発電機）

x = np.arange(len(test_k1))
width = 0.35

bars1 = ax2.bar(x - width/2, test_yco, width, label='Yco (Short-circuit capacity [p.u.])',
                color='steelblue', alpha=0.8)
bars2 = ax2.bar(x + width/2, [pgo_val]*4, width, label='Pgo per generator [p.u.]',
                color='coral', alpha=0.8)

ax2.set_xlabel('Test Case (k1 cycles, Bus No.)', fontsize=11)
ax2.set_ylabel('Parameter Value [p.u.]', fontsize=11)
ax2.set_title('.pop Parameter Rewrite Verification\n(4 test cases)', fontsize=11)
ax2.set_xticks(x)
ax2.set_xticklabels([f'k1={k}cyc\nBus{b}' for k, b in zip(test_k1, test_bus)], fontsize=9)
ax2.legend(fontsize=9)
ax2.set_ylim(0, 0.65)

# 値ラベル
for bar in bars1:
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
             f'{bar.get_height():.2f}', ha='center', va='bottom', fontsize=8)
for bar in bars2:
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
             f'{bar.get_height():.4f}', ha='center', va='bottom', fontsize=8)

plt.tight_layout()
out_path = '/home/ubuntu/pspt/CPATPub/main/Temp/dry_run_summary.png'
plt.savefig(out_path, dpi=150, bbox_inches='tight')
print(f'保存: {out_path}')
plt.close()
