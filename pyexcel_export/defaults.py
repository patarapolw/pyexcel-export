from datetime import datetime
from collections import OrderedDict
from io import BytesIO
import base64
import json


class Meta(OrderedDict):
    def __init__(self, **kwargs):
        default = OrderedDict([
            ('created', datetime.fromtimestamp(datetime.now().timestamp()).isoformat()),
            ('has_header', True),
            ('freeze_header', True),
            ('col_width_fit_param_keys', True),
            ('col_width_fit_ids', True),
            ('allow_table_hiding', True)
        ])

        default.update(**kwargs)

        super().__init__(**default)

    @property
    def excel_matrix(self):
        result = OrderedDict(self)

        for k, v in self.items():
            if type(v) not in (int, float, str):
                result[k] = {str(type(v)): v}

        return list(result.items())

    @property
    def view(self):
        result = OrderedDict(self)

        for k, v in self.items():
            if type(v) not in (int, float, str, bool):
                result[k] = {str(type(v)): v}

        return result

    @property
    def matrix(self):
        return list(self.view.items())

    def __setitem__(self, key, item):
        if isinstance(item, dict) and len(item) == 1:
            k, v = list(item.items())[0]

            if k == "<class 'bool'>":
                item = v
            elif k == "<class '_io.BytesIO'>":
                if isinstance(v, BytesIO):
                    item = v
                else:
                    item = BytesIO(base64.b64decode(v))

        super().__setitem__(key, item)

    def __repr__(self):
        output = OrderedDict()
        for k, v in self.items():
            output[k] = repr(v)

        return 'Meta({})'.format(json.dumps(output, indent=2))
