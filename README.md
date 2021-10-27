# Quant_Creon_API
주식 및 ETF 자동 매매를 위한 API입니다.
"파이썬 증권 데이터 분석(http://www.yes24.com/Product/Goods/90578506)" 책의 코드를 기반으로 일부 함수를 추가하였습니다

## AutoConnect.py
위 도서에 수록된 코드입니다. 대신증권 크레온 거래 프로그램에 자동 로그인이 가능합니다. 자세한 내용은 책을 참조 바랍니다.

## Default_API.py
위 도서에 수록된 코드에 일부 함수를 추가하였습니다. 책에 수록된 함수는 책을 참조 바랍니다.
책과 상이한 함수나 추가된 함수는 아래를 참고 하십시오.

- get_stock_balance(code, verbose=True)
```
verbose option을 통해 slack messaging이 컨트롤 가능합니다.
```
- get_stock_status(code)
```
해당 종목 코드의 종목명(str), 보유 수량(int), 수익률(double)을 반환합니다.
```
- get_stock_list()
```
현재 보유 중인 종목 리스트를 have_stock_list.json 파일로 저장하고 반환합니다.
```
- stock_monitor(code_list, upper_yield_limit, lower_yield_limit)
```
code_list: 모니터할 종목 리스트
upper_yield_limit: 수익률 상한 (e.g. 50 -> +50%)
lower_yield_limit: 수익률 하한 (e.g. 50 -> -50%)
상한/하한 수익률을 돌파한 종목 리스트를 반환합니다.
```
- buy_stock_list(code_list)
```
code_list의 종목을 모두 매수하여 동일 가중 포트폴리오를 구성합니다.
```
- sell_stock(code)
```
해당 종목을 매도합니다.
함수 내부의 일부 parameter를 수정하여 IOC/FOK 보통가/시장가 등의 option이 선택가능합니다.
```
- sell_stock_list(code_list)
```
code_list의 모든 종목을 매도합니다.
```

