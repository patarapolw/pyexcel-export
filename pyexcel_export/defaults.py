from collections import OrderedDict
from io import BytesIO
import base64
import binascii


class Meta(OrderedDict):
    @property
    def excel_matrix(self):
        result = OrderedDict(self)

        for k, v in self.items():
            assigned = False

            if self.get('bool_as_string', False):
                if v is True:
                    result[k] = 'true'
                    assigned = True
                elif v is False:
                    result[k] = 'false'
                    assigned = True

            # if not assigned:
            #     if type(v) not in (int, float, str, bool):
            #         result[k] = {str(type(v)): v}

        return list(result.items())

    @property
    def view(self):
        result = OrderedDict(self)

        # for k, v in self.items():
        #     if type(v) not in (int, float, str, bool, BytesIO):
        #         result[k] = {str(type(v)): v}

        return result

    @property
    def matrix(self):
        return [list(k_v_pair) for k_v_pair in self.view.items()]

    def __setitem__(self, key, item):
        if self.get('bool_as_string', False):
            if item in ('true', '\'true'):
                item = True
            elif item in ('false', '\'false'):
                item = False

        if isinstance(item, dict) and len(item) == 1:
            k, v = list(item.items())[0]

            if k == "<class 'bool'>":
                item = v
            elif k == "<class '_io.BytesIO'>":
                if isinstance(v, BytesIO):
                    item = v
                else:
                    item = BytesIO(base64.b64decode(v))

        if isinstance(item, str):
            try:
                item = BytesIO(base64.b64decode(item))
            except binascii.Error:
                pass

        super().__setitem__(key, item)

    def __repr__(self):
        output = []
        for k, v in self.items():
            output.append('(\'{}\', {})'.format(k, repr(v)))

        return 'Meta([\n  {}\n])'.format(',\n  '.join(output))
