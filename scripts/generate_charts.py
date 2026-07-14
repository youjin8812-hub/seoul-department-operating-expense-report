"""
집계·차트·분석 JSON 생성 (통합 스크립트)

기존 3개 스크립트를 통합했다:
- generate_visualizations.py  (monthly_trend.png 생성부만 흡수, dashboard.html 생성부는 archive)
- generate_visuals_enhanced.py
- generate_charts_extra_enhanced.py

산출물:
- output/charts/*.png (8종)
- workspace/analysis_enhanced.json

집계 기준: 해당년도 -> 해당월 -> 전체부서명 -> 비목 -> 집행금액 합계
개인정보(전화번호, 작성자, 집행대상 원문)는 어떤 출력에도 포함하지 않는다.
고액 집행은 "고액 집행 후보"로만 표현하며 오류/부적정으로 단정하지 않는다.
"""
import argparse
import json
import os

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from adjustText import adjust_text

CHARTS_DIR = "output/charts"
JSON_PATH = "workspace/analysis_enhanced.json"

plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

CAT_COLOR = {"기관": "#1f4e8c", "부서": "#c98a1f", "시책": "#3f9142", "정원": "#a13f6b"}
STACK_COLORS = ["#1f4e8c", "#3f72af", "#6a97cf", "#9dbbe0", "#c9d6e3", "#b0b0b0"]

REQUIRED_COLUMNS = ["해당년도", "해당월", "전체부서명", "비목", "집행금액"]


def load_data(csv_path):
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing} (원본 컬럼명을 임의로 변경하지 않았는지 확인)")
    df["집행금액"] = df["집행금액"].astype("int64")
    df["해당월"] = df["해당월"].astype(int)
    return df


def sanity_check(df):
    """데이터셋에 무관한 구조적 검증만 수행한다(특정 데이터셋의 고정값을 가정하지 않음)."""
    assert len(df) > 0, "데이터가 비어 있음"
    assert (df["집행금액"] >= 0).all(), "음수 집행금액 존재"
    monthly_sum = df.groupby("해당월")["집행금액"].sum().sum()
    assert monthly_sum == df["집행금액"].sum(), "월별 합계와 전체 합계 불일치"
    category_sum = df.groupby("비목")["집행금액"].sum().sum()
    assert category_sum == df["집행금액"].sum(), "비목별 합계와 전체 합계 불일치"


# ---------- 월별·부서별 집행 현황 (구 monthly_trend) ----------
def monthly_by_dept_top5(df, top_n=5):
    dept_totals = df.groupby("전체부서명")["집행금액"].sum().sort_values(ascending=False)
    top_depts = list(dept_totals.head(top_n).index)
    df = df.copy()
    df["부서구분"] = df["전체부서명"].where(df["전체부서명"].isin(top_depts), "기타")
    pivot = df.pivot_table(index="해당월", columns="부서구분", values="집행금액",
                             aggfunc="sum", fill_value=0).sort_index()
    ordered_cols = [c for c in top_depts + ["기타"] if c in pivot.columns]
    return pivot[ordered_cols]


def chart_monthly_trend(monthly_dept, period_label):
    months = monthly_dept.index
    totals = monthly_dept.sum(axis=1)
    fig, ax = plt.subplots(figsize=(7.5, 4.5), dpi=300)
    bottom = pd.Series(0, index=months, dtype="float64")
    for col, color in zip(monthly_dept.columns, STACK_COLORS):
        vals = monthly_dept[col] / 1_000_000
        ax.bar([f"{m}월" for m in months], vals, bottom=bottom / 1_000_000, color=color, label=col)
        bottom += monthly_dept[col]
    ax.set_ylabel("집행액(백만원)")
    ax.set_title(f"서울시 본청 업무추진비 월별·부서별 집행 현황 ({period_label})", fontsize=12, fontweight="bold")
    ax.set_ylim(0, totals.max() / 1_000_000 * 1.25)
    for i, m in enumerate(months):
        ax.text(i, totals[m] / 1_000_000 + 5, f"{totals[m]/1_000_000:,.0f}", ha="center", fontsize=9, fontweight="bold")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=3, fontsize=7.5, frameon=False)
    fig.text(0.02, 0.01, "집계 기준: 해당월×전체부서명×집행금액 합계(상위 5개 부서 개별 표시, 나머지는 기타) | 원본 CSV 기준",
              fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.05, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/monthly_trend.png")
    plt.close(fig)


# ---------- 반복 집행 패턴 ----------
def analyze_repeat_patterns(df, top_n=10):
    grp = df.groupby(["전체부서명", "비목", "집행금액"])
    agg = grp.agg(반복횟수=("해당월", "size"), 반복월목록=("해당월", lambda s: sorted(s.unique().tolist())))
    agg = agg.reset_index()
    agg["반복월수"] = agg["반복월목록"].map(len)
    agg = agg[agg["반복횟수"] >= 2].sort_values("반복횟수", ascending=False).head(top_n)
    return agg


def chart_repeat_patterns(agg):
    fig, ax = plt.subplots(figsize=(8, 5), dpi=300)
    y_pos = range(len(agg))
    bar_colors = [CAT_COLOR.get(c, "#888888") for c in agg["비목"]]
    ax.barh(y_pos, agg["반복횟수"], color=bar_colors)
    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(agg["전체부서명"], fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("반복 횟수")
    ax.set_title("반복 집행 패턴 상위 10개 (동일 부서·비목·금액 조합)", fontsize=11, fontweight="bold")
    for i, (_, row) in enumerate(agg.iterrows()):
        ax.text(row["반복횟수"] + 0.15, i, f"{row['비목']} · {row['집행금액']:,}원 · {row['반복월수']}개월 분산",
                 va="center", fontsize=7.5, color="#333333")
    if len(agg) > 0:
        ax.set_xlim(0, agg["반복횟수"].max() * 1.9)
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in CAT_COLOR.values()]
    ax.legend(handles, CAT_COLOR.keys(), title="비목", loc="lower right", fontsize=7.5, title_fontsize=8)
    fig.text(0.02, 0.01,
              "집계 기준: 전체부서명×비목×집행금액 동일 조합 반복 횟수(2회 이상) | "
              "주의: 반복 사유·정례성은 데이터로 확인되지 않음, 동일 금액 반복 사실만 표시",
              fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.04, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/repeat_patterns_enhanced.png")
    plt.close(fig)


# ---------- 월별 비목 구성 변화 ----------
def analyze_monthly_category_mix(df):
    pivot = df.pivot_table(index="해당월", columns="비목", values="집행금액", aggfunc="sum", fill_value=0).sort_index()
    share = pivot.div(pivot.sum(axis=1), axis=0) * 100
    return pivot, share


def chart_monthly_category_mix(share, period_label):
    fig, ax = plt.subplots(figsize=(7.5, 4.5), dpi=300)
    bottom = pd.Series(0, index=share.index, dtype="float64")
    months = [f"{m}월" for m in share.index]
    for col in share.columns:
        ax.bar(months, share[col], bottom=bottom, color=CAT_COLOR.get(col, "#888888"), label=col)
        bottom += share[col]
    ax.set_ylabel("비중(%)")
    ax.set_ylim(0, 100)
    ax.set_title(f"월별 비목 구성 변화 ({period_label})", fontsize=12, fontweight="bold")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.12), ncol=4, fontsize=8.5, frameon=False)
    fig.text(0.02, 0.01, "집계 기준: 해당월×비목×집행금액 비중(100% 환산)", fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.05, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/monthly_category_mix_enhanced.png")
    plt.close(fig)


# ---------- 전월 대비 증감 ----------
def analyze_mom_change(df):
    monthly = df.groupby("해당월").agg(총액=("집행금액", "sum"), 건수=("집행금액", "size")).sort_index()
    monthly["건당평균"] = monthly["총액"] / monthly["건수"]
    monthly["전월대비증감액"] = monthly["총액"].diff()
    monthly["전월대비증감률(%)"] = monthly["총액"].pct_change() * 100
    monthly["전월대비건수증감"] = monthly["건수"].diff()
    monthly["전월대비건당평균증감"] = monthly["건당평균"].diff()
    return monthly


def chart_mom_change(monthly):
    fig, ax = plt.subplots(figsize=(7, 4), dpi=300)
    months = [f"{m}월" for m in monthly.index]
    rates = monthly["전월대비증감률(%)"].values
    colors = ["#bbbbbb" if np.isnan(r) else ("#1f4e8c" if r >= 0 else "#b23b3b") for r in rates]
    bars = ax.bar(months, np.nan_to_num(rates, nan=0.0), color=colors)
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_ylabel("전월 대비 증감률(%)")
    ax.set_title("월별 집행액 전월 대비 증감률 (첫 달은 비교 불가)", fontsize=11, fontweight="bold")
    for bar, r in zip(bars, rates):
        label = "비교 불가" if np.isnan(r) else f"{r:+.1f}%"
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + (3 if bar.get_height() >= 0 else -8),
                 label, ha="center", fontsize=8)
    fig.text(0.02, 0.01, "집계 기준: 해당월×집행금액 합계의 전월 대비 변화율", fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/monthly_mom_change_enhanced.png")
    plt.close(fig)


# ---------- 동일기간 부서 구조 ----------
def analyze_dept_structure(df):
    n_months = df["해당월"].nunique()
    coverage = df.groupby("전체부서명")["해당월"].nunique()
    full_depts = coverage[coverage == n_months].index.tolist()
    sub = df[df["전체부서명"].isin(full_depts)]

    dept_summary = sub.groupby("전체부서명").agg(총액=("집행금액", "sum"), 건수=("집행금액", "size"))
    dept_summary["건당평균"] = dept_summary["총액"] / dept_summary["건수"]

    monthly_by_dept = sub.pivot_table(index="전체부서명", columns="해당월", values="집행금액", aggfunc="sum", fill_value=0)
    dept_summary["월간표준편차"] = monthly_by_dept.std(axis=1)
    dept_summary["월간평균"] = monthly_by_dept.mean(axis=1)
    dept_summary["변동계수"] = dept_summary["월간표준편차"] / dept_summary["월간평균"]

    cat_pivot = sub.pivot_table(index="전체부서명", columns="비목", values="집행금액", aggfunc="sum", fill_value=0)
    cat_share = cat_pivot.div(cat_pivot.sum(axis=1), axis=0) * 100

    dept_summary = dept_summary.sort_values("총액", ascending=False)
    if len(cat_share) > 0:
        max_share = cat_share.max(axis=1)
        dominant_cat = cat_share.idxmax(axis=1)
        dept_summary["최대비목비중(%)"] = max_share.reindex(dept_summary.index)
        dept_summary["최대비목"] = dominant_cat.reindex(dept_summary.index)
        dept_summary["비목의존도높음(>=60%)"] = dept_summary["최대비목비중(%)"] >= 60

    return dept_summary, cat_share.reindex(dept_summary.index), len(full_depts), df["전체부서명"].nunique()


def chart_dept_structure_heatmap(cat_share):
    fig, ax = plt.subplots(figsize=(7, 6), dpi=300)
    data = cat_share.values
    im = ax.imshow(data, cmap="Blues", aspect="auto", vmin=0, vmax=100)
    ax.set_xticks(range(len(cat_share.columns)))
    ax.set_xticklabels(cat_share.columns, fontsize=9)
    ax.set_yticks(range(len(cat_share.index)))
    ax.set_yticklabels(cat_share.index, fontsize=7.5)
    for i in range(data.shape[0]):
        for j in range(data.shape[1]):
            val = data[i, j]
            color = "white" if val > 55 else "black"
            ax.text(j, i, f"{val:.0f}%", ha="center", va="center", fontsize=7, color=color)
    ax.set_title("동일기간(전 기간 보유) 부서 × 비목 구성비", fontsize=11, fontweight="bold")
    fig.colorbar(im, ax=ax, label="비중(%)", shrink=0.7)
    fig.text(0.02, 0.01, "집계 기준: 전체부서명(전기간 보유)×비목×집행금액 비중", fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/dept_structure_heatmap_enhanced.png")
    plt.close(fig)


def chart_dept_scatter(df):
    n_months = df["해당월"].nunique()
    coverage = df.groupby("전체부서명")["해당월"].nunique()
    full_depts = coverage[coverage == n_months].index.tolist()
    sub = df[df["전체부서명"].isin(full_depts)]

    grp = sub.groupby("전체부서명").agg(총액=("집행금액", "sum"), 건수=("집행금액", "size"))
    grp["건당평균"] = grp["총액"] / grp["건수"]
    cat_pivot = sub.pivot_table(index="전체부서명", columns="비목", values="집행금액", aggfunc="sum", fill_value=0)
    dom_cat = cat_pivot.idxmax(axis=1)
    grp["주비목"] = dom_cat.reindex(grp.index)

    fig, ax = plt.subplots(figsize=(7.5, 5), dpi=300)
    sizes = (grp["건당평균"] / grp["건당평균"].max()) * 800 + 80
    colors = [CAT_COLOR.get(c, "#888888") for c in grp["주비목"]]
    ax.scatter(grp["건수"], grp["총액"] / 1_000_000, s=sizes, c=colors, alpha=0.75, edgecolors="white", linewidth=0.8)
    texts = [ax.text(row["건수"], row["총액"] / 1_000_000, name, fontsize=7) for name, row in grp.iterrows()]
    adjust_text(texts, ax=ax, arrowprops=dict(arrowstyle="-", color="gray", lw=0.6), expand_points=(1.4, 1.4))
    ax.set_xlabel("집행 건수")
    ax.set_ylabel("집행액(백만원)")
    ax.set_title("동일기간(전 기간 보유) 부서 — 규모와 활동성 (점 크기=건당평균, 색=주비목)", fontsize=11, fontweight="bold")
    handles = [plt.Line2D([0], [0], marker='o', color='w', markerfacecolor=c, markersize=8, label=k) for k, c in CAT_COLOR.items()]
    ax.legend(handles=handles, title="주비목", loc="upper left", fontsize=8, title_fontsize=8.5)
    fig.text(0.02, 0.01, "집계 기준: 전체부서명(전기간 보유)×집행금액", fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/dept_scatter_enhanced.png")
    plt.close(fig)


# ---------- 고액 집행 후보 및 집중도 ----------
def analyze_high_value_and_concentration(df):
    amt = df["집행금액"]
    dist = {"중앙값": float(amt.median()), "평균": float(amt.mean()), "p75": float(amt.quantile(0.75)),
            "p90": float(amt.quantile(0.90)), "p95": float(amt.quantile(0.95))}
    q1, q3 = amt.quantile(0.25), amt.quantile(0.75)
    iqr_fence = q3 + 1.5 * (q3 - q1)
    iqr_candidates = df[amt > iqr_fence]
    p95_threshold = amt.quantile(0.95)
    p95_candidates = df[amt >= p95_threshold]

    cat_dist = df.groupby("비목")["집행금액"].describe()[["50%", "mean", "max"]]

    dept_totals = df.groupby("전체부서명")["집행금액"].sum().sort_values(ascending=False)
    total = dept_totals.sum()
    shares = dept_totals / total
    top5_share = float(shares.head(5).sum() * 100)
    top10_share = float(shares.head(10).sum() * 100)
    hhi = float((shares ** 2).sum() * 10000)

    result = {
        "distribution": dist, "iqr_fence": float(iqr_fence),
        "iqr_candidate_count": int(len(iqr_candidates)), "p95_threshold": float(p95_threshold),
        "p95_candidate_count": int(len(p95_candidates)), "category_distribution": cat_dist.to_dict(orient="index"),
        "top5_share_pct": top5_share, "top10_share_pct": top10_share, "hhi": hhi,
        "dept_totals_desc": dept_totals,
        "iqr_candidates_by_month": iqr_candidates.groupby("해당월").size().to_dict(),
        "iqr_candidates_by_category": iqr_candidates.groupby("비목").size().to_dict(),
    }
    return result


def chart_amount_distribution(df, iqr_fence, p95):
    amt = df["집행금액"]
    fig, ax = plt.subplots(figsize=(7.5, 4.5), dpi=300)
    ax.hist(amt, bins=40, color="#3f72af", edgecolor="white", linewidth=0.3)
    ax.axvline(iqr_fence, color="#e08a00", linestyle="--", linewidth=1.3)
    ax.axvline(p95, color="#b23b3b", linestyle="--", linewidth=1.3)
    ax.text(iqr_fence, ax.get_ylim()[1] * 0.9, f" IQR상단경계\n {iqr_fence:,.0f}원", color="#e08a00", fontsize=7.5)
    ax.text(p95, ax.get_ylim()[1] * 0.7, f" 상위5%기준\n {p95:,.0f}원", color="#b23b3b", fontsize=7.5)
    ax.set_xlabel("집행금액(원)")
    ax.set_ylabel("건수")
    ax.set_title("전체 집행금액 분포 및 고액 집행 후보 기준선", fontsize=11, fontweight="bold")
    fig.text(0.02, 0.01, f"집계 기준: 집행금액(전체 {len(df):,}건) | 고액 집행 후보는 오류·부적정 집행으로 단정하지 않음",
              fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.03, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/amount_distribution_enhanced.png")
    plt.close(fig)


def chart_concentration_pareto(dept_totals_desc, top_n=15):
    top = dept_totals_desc.head(top_n)
    total = dept_totals_desc.sum()
    cum_share = top.cumsum() / total * 100

    fig, ax1 = plt.subplots(figsize=(8, 4.5), dpi=300)
    ax1.bar(range(len(top)), top.values / 1_000_000, color="#1f4e8c")
    ax1.set_xticks(range(len(top)))
    ax1.set_xticklabels([f"{i+1}" for i in range(len(top))], fontsize=8)
    ax1.set_xlabel("부서 순위 (집행액 기준)")
    ax1.set_ylabel("집행액(백만원)")
    ax1.set_title(f"부서별 집행액 집중도 — 파레토 차트 (상위 {len(top)}개 부서)", fontsize=11, fontweight="bold")

    ax2 = ax1.twinx()
    ax2.plot(range(len(top)), cum_share.values, color="#b23b3b", marker="o", markersize=4)
    ax2.set_ylabel("누적 점유율(%)", color="#b23b3b")
    ax2.set_ylim(0, 105)
    ax2.tick_params(axis="y", labelcolor="#b23b3b")

    for idx in [4, 9]:
        if idx < len(cum_share):
            ax2.annotate(f"상위{idx+1}: {cum_share.iloc[idx]:.1f}%", (idx, cum_share.iloc[idx]),
                         textcoords="offset points", xytext=(0, 10), fontsize=8, color="#b23b3b", ha="center")

    fig.text(0.02, 0.01, "집계 기준: 전체부서명×집행금액 합계, 순위 기준 누적 점유율 | "
              "주의: 집중도가 높다는 사실 자체를 부정적으로 해석하지 않음", fontsize=6.5, color="gray")
    fig.tight_layout(rect=[0, 0.04, 1, 1])
    fig.savefig(f"{CHARTS_DIR}/concentration_pareto_enhanced.png")
    plt.close(fig)


def run(csv_path):
    os.makedirs(CHARTS_DIR, exist_ok=True)
    os.makedirs("workspace", exist_ok=True)

    df = load_data(csv_path)
    sanity_check(df)

    years = sorted(df["해당년도"].unique().tolist())
    months = sorted(df["해당월"].unique().tolist())
    period_label = f"{years[0]}년 {months[0]}~{months[-1]}월" if len(years) == 1 else f"{years[0]}~{years[-1]}년"

    monthly_dept = monthly_by_dept_top5(df)
    chart_monthly_trend(monthly_dept, period_label)

    repeat_agg = analyze_repeat_patterns(df)
    chart_repeat_patterns(repeat_agg)

    cat_pivot, cat_share = analyze_monthly_category_mix(df)
    chart_monthly_category_mix(cat_share, period_label)

    mom = analyze_mom_change(df)
    chart_mom_change(mom)

    dept_structure, dept_cat_share, n_full_depts, n_total_depts = analyze_dept_structure(df)
    if len(dept_cat_share) > 0:
        chart_dept_structure_heatmap(dept_cat_share)
        chart_dept_scatter(df)

    hv = analyze_high_value_and_concentration(df)
    chart_amount_distribution(df, hv["iqr_fence"], hv["distribution"]["p95"])
    chart_concentration_pareto(hv["dept_totals_desc"])

    total_check = int(df["집행금액"].sum())
    output = {
        "meta": {
            "총행수": len(df), "총집행액": total_check, "분석기간": period_label,
            "동일기간비교가능부서수": n_full_depts, "전체부서수": n_total_depts,
            "집계기준": "해당년도->해당월->전체부서명->비목->집행금액 합계",
        },
        "1_repeat_patterns": [
            {"전체부서명": r["전체부서명"], "비목": r["비목"], "집행금액": int(r["집행금액"]),
             "반복횟수": int(r["반복횟수"]), "반복월수": int(r["반복월수"]), "반복월목록": r["반복월목록"]}
            for _, r in repeat_agg.iterrows()
        ],
        "2_monthly_category_mix": {
            "amount": cat_pivot.astype(int).to_dict(orient="index"),
            "share_pct": cat_share.round(2).to_dict(orient="index"),
        },
        "3_mom_change": mom.round(2).replace({np.nan: None}).to_dict(orient="index"),
        "4_dept_structure_same_period": dept_structure.round(2).replace({np.nan: None}).to_dict(orient="index"),
        "4_dept_category_share_pct": dept_cat_share.round(2).to_dict(orient="index"),
        "5_high_value_and_concentration": {k: v for k, v in hv.items() if k != "dept_totals_desc"},
    }
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"[generate_charts] 완료: 8개 차트 + {JSON_PATH}")
    print(f"  총 집행액={total_check:,}원 | 총건수={len(df):,}건 | "
          f"동일기간비교가능부서={n_full_depts}/{n_total_depts}")
    return output


def main():
    parser = argparse.ArgumentParser(description="CSV 집계·차트·분석 JSON 생성")
    parser.add_argument("--input", default="input/seoul_expenses.csv", help="원본 CSV 경로")
    args = parser.parse_args()
    run(args.input)


if __name__ == "__main__":
    main()
