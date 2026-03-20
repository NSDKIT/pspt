import pandas as pd
from itertools import combinations_with_replacement

def generate_combinations(start, end, group_size, i):
    # 数字の範囲を定義
    numbers = range(start, end + 1)

    # 重複組み合わせを生成
    combinations = list(combinations_with_replacement(numbers, i))

    # 組み合わせをグループに分ける
    groups = [combinations[j:j + group_size] for j in range(0, len(combinations), group_size)]

    # 最後のグループを補完するために、最初のグループから必要な数を追加
    if len(groups[-1]) < group_size:
        needed = group_size - len(groups[-1])
        groups[-1].extend(groups[0][:needed])

    # 各グループをDataFrameに変換
    df_list = []
    for idx, group in enumerate(groups):
        df = pd.DataFrame(group, columns=[f'Element{k+1}' for k in range(i)])
        df_list.append(df)

    return df_list

df_list = generate_combinations(11, 47, 10, 1)
print(len(df_list))