import pandas as pd

# 예제용 DataFrame 생성
data = {'A': ['a', 'b', 'c', 'd'],
        'B': [1, 2, 3, 4]}
df = pd.DataFrame(data)

# 특정 조건을 만족하는 행 선택하여 복사
df_copy = df[df['A'].isin(['a', 'b'])].copy()

# 선택된 행을 기존 DataFrame에 추가
df = pd.concat([df, df_copy], ignore_index=True)

# 결과 출력
print(df)