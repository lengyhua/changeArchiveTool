SUCCESS = '成功'
FAIL = '失败'
NOT_FOUND = 'NOT FOUND'
ERROR_INFO = "填写信息有误"


class DeleteInfo:
    def __init__(self, row, value, result_col=2, result=SUCCESS):
        self.row = row
        self.value = value
        self.result_row = self.row
        self.result_col = result_col
        self.result = result


class UpdateInfo:
    def __init__(self, row, values, result_col=10, result=SUCCESS):
        self.row = row
        self.values = values
        self.result_row = self.row
        self.result_col = result_col
        self.result = result


class MergeInfo:
    def __init__(self, row, from_archive, to_archive, result_col=3, result=SUCCESS):
        self.row = row
        self.from_archive = from_archive
        self.to_archive = to_archive
        self.result_row = row
        self.result_col = result_col
        self.result = result
