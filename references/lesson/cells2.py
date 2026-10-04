# 2차시 칸 목록 — 노트북과 슬라이드 코드 카드가 모두 여기서 나온다 (한 곳에서만 고친다)
# 예제용 가상 수업이다. 실제 수업에서는 이 파일을 차시마다 하나씩 둔다.
TITLE = "데이터 첫걸음 · 2차시 실습 — 배열로 한꺼번에 계산하기"

CELLS = [
  {"section": "① 배열 만들기"},
  {"no": 1, "title": "일주일 판매량을 배열로", "code": """
import numpy as np

# 월요일부터 일요일까지
sales = np.array([12, 15, 9, 20, 18, 25, 30])
print(sales * 2)   # 칸마다 두 배
"""},
  {"no": 2, "title": "평균과 가장 많이 판 날", "code": """
print(sales.mean().round(1))
print(sales.argmax())   # 0 = 월요일
"""},
  {"section": "② 그림으로 보기"},
  {"no": 3, "title": "꺾은선으로 그려 보기", "code": """
days = ['월', '화', '수', '목', '금', '토', '일']
plt.plot(days, sales, marker='o')
plt.ylabel('판매량 (개)')
plt.show()
"""},
]
