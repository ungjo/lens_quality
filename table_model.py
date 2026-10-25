"""DataFrame을 PyQt6 표에 표시하는 모델입니다. 파일을 읽거나 저장하지 않습니다.

gui.py에서 사용할 연결 예:
    model = DataFrameTableModel(report.result, report.detail)
    proxy = LensFilterProxyModel()
    proxy.setSourceModel(model)
    table_view.setModel(proxy)
    table_view.setSortingEnabled(True)

모델 생성과 set_dataframe()/set_filters() 호출은 GUI 스레드에서 수행하세요.
검사 작업자는 DataFrame을 전달하고, GUI가 그 결과를 받아 모델을 갱신합니다.
"""

from numbers import Real #값이 숫자인지 확인

import pandas as pd
from pandas.api.types import is_scalar #DataFrame과 누락값 처리
# QAbstractTableModel : DataFrame을 GUI표에 연결, QModelIndex : 표의 특정 셀 위치를 나타냄, QSortFilterProxyModel : 정렬/필터 기능의 기본 틀, Qt : 표시값, 배경색, 정렬 방향 등의 구분값
from PyQt6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt
from PyQt6.QtGui import QBrush, QColor # 셀의 배경색과 글자색 표현


PASS_COLOR = "#D9EAD3"
FAIL_COLOR = "#F4CCCC"
FAILED_VALUE_COLOR = "#D9D9D9"
COLORED_TEXT_COLOR = "#1F2937"

# 셀 값을 사용하기 편하게 하기 위한 함수
def _plain_value(value):
    """누락값은 None, NumPy 숫자는 Python 숫자로 바꿉니다."""
    if is_scalar(value) and pd.isna(value): #누락값은 None
        return None
    if hasattr(value, "item"):
        return value.item() # item() : 값을 일반 Python값(실수, 정수, 복소수)으로 바꿈.
    return value

# 행/열 개수, 컬럼 이름, 셀 값과 색상을 GUI에 전달 -> 데이터프레임을 표시하기 위해서
class DataFrameTableModel(QAbstractTableModel):
    """네 DB 테이블과 검사 결과 DataFrame을 읽기 전용 표로 보여줍니다.

    details는 결과표의 불합격 원인 셀을 찾을 때 사용합니다.
    원본 DataFrame의 순서와 값을 보호하기 위해 복사본을 보관합니다.
    """

    RAW_VALUE_ROLE = int(Qt.ItemDataRole.UserRole) + 1
    PRODUCT_ID_ROLE = int(Qt.ItemDataRole.UserRole) + 2 #해당 행의 제품번호

    def __init__(self, dataframe=None, details=None, parent=None):
        super().__init__(parent) #PyQt6 기본 모델의 초기화 작업 실행
        self._frame = pd.DataFrame() #화면에 표시할 DataFrame 보관
        self._columns = {} #컬럼 이름과 위치의 대응 관계 보관
        self._failed_cells = set() #회색으로 표시할 제품번호/검사항목 조합 보관
        self.set_dataframe(dataframe, details) #전달받은 데이터로 표 설정

    def set_dataframe(self, dataframe=None, details=None):
        """표 데이터를 교체하고 이전 불합격 색상 정보를 함께 초기화합니다."""
        frame = pd.DataFrame() if dataframe is None else dataframe.copy(deep=True) #데이터가 없으면 빈 표를 만들고 데이터가 있으면 복사해서 보관
        if not frame.columns.is_unique or not all(isinstance(name, str) for name in frame.columns):
            raise ValueError("표의 컬럼 이름은 중복되지 않는 문자열이어야 합니다.")

        failed_cells = set()
        if details is not None and not details.empty:
            required = ["제품번호", "검사항목", "항목판정"]
            missing = [name for name in required if name not in details.columns]
            if missing or not details.columns.is_unique:
                raise ValueError("항목별 상세에는 중복 없는 제품번호, 검사항목, 항목판정 컬럼이 필요합니다.")
            for product_id, item, status in details[required].itertuples(index=False, name=None):
                product_id, item, status = map(_plain_value, (product_id, item, status))
                if status == "불합격" and product_id is not None and item is not None:
                    failed_cells.add((str(product_id), str(item))) #상세 결과에서 불합격한 항목을 찾으면 기록

        # 연결된 표와 프록시에도 데이터 전체가 바뀌었음을 알립니다.
        self.beginResetModel()
        self._frame = frame
        self._columns = {name: position for position, name in enumerate(frame.columns)}
        self._failed_cells = failed_cells
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()): #행 개수 반환
        return 0 if parent.isValid() else len(self._frame)

    def columnCount(self, parent=QModelIndex()): #열 개수 반환
        return 0 if parent.isValid() else len(self._frame.columns)

    def column_index(self, name): #컬럼 이름으로 열 위치 찾기
        """열 이름에 해당하는 위치를 반환합니다. 없는 열은 None입니다."""
        return self._columns.get(name)

    def value_at(self, row, column): #지정한 행/열의 셀 값 반환
        """DataFrame의 인덱스 라벨과 관계없이 위치로 원본 값을 가져옵니다."""
        return _plain_value(self._frame.iat[row, column]) #.iat : 행과 열 위치의 셀 선택. .iat[0,0] : 첫번째 행 첫번째 열

    def record_at(self, row): #한 행 전체를 사전 형태로 반환
        """원본 표의 한 행을 사전으로 반환합니다. 정렬된 표는 프록시의 메서드를 사용하세요."""
        if not 0 <= row < self.rowCount():
            raise IndexError("표의 행 범위를 벗어났습니다.")
        return {name: self.value_at(row, column) for name, column in self._columns.items()}

    def _background_color(self, row, column, value):
        name = self._frame.columns[column]
        if name in ("판정", "항목판정"):
            return {"합격": PASS_COLOR, "불합격": FAIL_COLOR}.get(value)

        # 상세표는 해당 행의 항목판정으로 측정값의 색상을 결정합니다.
        status_column = self.column_index("항목판정")
        if name == "측정값" and status_column is not None:
            if self.value_at(row, status_column) == "불합격":
                return FAILED_VALUE_COLOR

        # 결과표는 제품번호와 검사항목으로 원인을 찾습니다.
        product_column = self.column_index("제품번호")
        if "판정" in self._columns and product_column is not None:
            product_id = self.value_at(row, product_column)
            if product_id is not None and (str(product_id), name) in self._failed_cells:
                return FAILED_VALUE_COLOR
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole): #DisplayRole : 화면에 표시할 글자 반환
        if not index.isValid() or index.model() is not self:
            return None
        row, column = index.row(), index.column()
        if not (0 <= row < self.rowCount() and 0 <= column < self.columnCount()):
            return None
        value = self.value_at(row, column)

        if role == self.RAW_VALUE_ROLE:
            return value
        if role == self.PRODUCT_ID_ROLE:
            product_column = self.column_index("제품번호")
            return None if product_column is None else self.value_at(row, product_column)
        if role == Qt.ItemDataRole.DisplayRole:
            # 긴 판정 사유는 한 줄로, 전체 문장은 마우스 툴팁으로 표시합니다.
            return "" if value is None else str(value).replace("\r\n", "\n").replace("\n", " / ")
        if role == Qt.ItemDataRole.ToolTipRole: #ToolTipRole : 마우스를 올렸을 때 보여줄 전체 내용 반환
            return "" if value is None else str(value)
        if role == Qt.ItemDataRole.TextAlignmentRole: #TextAlignmentRole : 숫자는 오른쪽, 문자는 왼쪽 정렬
            horizontal = Qt.AlignmentFlag.AlignRight if isinstance(value, Real) else Qt.AlignmentFlag.AlignLeft
            return horizontal | Qt.AlignmentFlag.AlignVCenter
        if role in (Qt.ItemDataRole.BackgroundRole, Qt.ItemDataRole.ForegroundRole): #ForegroundRole : 글자색반환
            color = self._background_color(row, column, value)
            if color is not None:
                return QBrush(QColor(color if role == Qt.ItemDataRole.BackgroundRole else COLORED_TEXT_COLOR))
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole): #가로방향에는 컬럼 이름, 세로방향에는 행 번호 표시
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal and 0 <= section < self.columnCount():
            return self._frame.columns[section]
        if orientation == Qt.Orientation.Vertical and 0 <= section < self.rowCount():
            return str(section + 1)
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        # ItemIsEditable을 제외하여 표에서 원본 측정값을 수정하지 못하게 합니다.
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable  #셀 사용과 선택만 가능. 셀의 값 수정 불가

# 정렬, 검색,판정/소재 필터 -> 원본 데이터를 유지하면서 화면에 보이는 순서와 행만 바꾸기 위해서
class LensFilterProxyModel(QSortFilterProxyModel):
    """원본 데이터를 유지하면서 표시 순서와 보이는 행만 바꿉니다.

    숫자는 숫자 크기로 정렬하고, 오름차순에서 누락값을 뒤에 배치합니다.
    검색어는 모든 열에서 찾습니다. 판정과 소재 조건은 정확히 일치해야 하며,
    검색/판정/소재 조건을 모두 만족하는 행만 표시합니다.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._search_text = ""
        self._status = None
        self._material = None
        self.setDynamicSortFilter(True)
        self.setSortRole(DataFrameTableModel.RAW_VALUE_ROLE) #RAW_VALUE_ROLE : 정렬에 사용할 원래 값

    def setSourceModel(self, model): #setSourceModel() : 어떤 원본 모델을 사용할지 연결
        if model is not None and not isinstance(model, DataFrameTableModel):
            raise TypeError("DataFrameTableModel을 원본 모델로 지정하세요.")
        super().setSourceModel(model)

    def set_filters(self, text="", status=None, material=None): #set_filters() : 검색어/판정/소재 조건 설정
        """빈 값 또는 '전체'는 해당 선택 조건을 해제합니다.

        status는 판정 또는 항목판정 열에 적용하고 material은 소재종류 열에 적용합니다.
        조건에 필요한 열이 없는 표에는 해당 조건을 지정하지 마세요.
        """
        status = None if status in (None, "", "전체") else status
        material = None if material in (None, "", "전체") else material
        if status not in (None, "합격", "불합격", "판정불가"):
            raise ValueError("판정 조건은 전체, 합격, 불합격, 판정불가 중 하나여야 합니다.")
        search_text = str(text or "").strip().casefold()

        # PyQt6 6.11에서 제공하는 필터 변경 알림으로 화면을 갱신합니다.
        self.beginFilterChange()
        self._search_text = search_text
        self._status = status
        self._material = material
        self.endFilterChange(QSortFilterProxyModel.Direction.Rows)

    def filterAcceptsRow(self, source_row, source_parent): #filterAceptRow() : 각 행을 화면에 표시할지 결정
        model = self.sourceModel()
        if model is None:
            return False

        if self._status is not None:
            column = model.column_index("판정")
            if column is None:
                column = model.column_index("항목판정")
            if column is None or model.value_at(source_row, column) != self._status:
                return False
        if self._material is not None:
            column = model.column_index("소재종류")
            if column is None or model.value_at(source_row, column) != self._material:
                return False
        if self._search_text:
            for column in range(model.columnCount()):
                value = model.value_at(source_row, column)
                if value is not None and self._search_text in str(value).casefold():
                    return True
            return False
        return True

    def lessThan(self, left, right): #lessThan() : 두 값을 비교해 정렬 순서 결정
        def sort_key(value):
            if value is None:
                return (2, "")
            if isinstance(value, Real):
                return (0, value)
            return (1, str(value).casefold())

        model = self.sourceModel()
        return sort_key(model.data(left, self.sortRole())) < sort_key(model.data(right, self.sortRole()))

    def record_at(self, row): #record_at() : 정렬된 화면에서 선택한 행의 원본 데이터 반환
        """정렬/검색 후 화면에서 선택한 행을 원본 행으로 연결해 반환합니다."""
        model = self.sourceModel()
        if model is None or not 0 <= row < self.rowCount():
            raise IndexError("표의 행 범위를 벗어났습니다.")
        source_index = self.mapToSource(self.index(row, 0)) #현재 화면의 이 행이 원본 표에서는 어느 행인지 찾기. 정렬 후 화면 첫번째 행이 바뀌었다면 그 행을 선택했을 때 해당 인덱스번호의 정보를 가져옮.
        if not source_index.isValid():
            raise IndexError("선택한 행에 표시할 컬럼이 없습니다.")
        return model.record_at(source_index.row())
