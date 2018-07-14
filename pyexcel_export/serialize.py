import uuid
import json
import base64
from io import BytesIO


class RowExport(object):
    def __init__(self, raw_row):
        self.value = []
        for raw_cell in raw_row:
            self.value.append(json.dumps(raw_cell, ensure_ascii=False, cls=MyEncoder))

    def __repr__(self):
        if not isinstance(self.value, list):
            return repr(self.value)
        else:  # Sort the representation of any dicts in the list.
            reps = ('{{{}}}'.format(', '.join(
                        ('{!r}:{}'.format(k, v) for k, v in sorted(v.items()))
                    )) if isinstance(v, dict)
                        else
                    repr(v) for v in self.value)
            return '[' + ', '.join(reps) + ']'

    @property
    def data(self):
        raw_row = []
        for cell in self.value:
            raw_row.append(json.loads(cell))

        return raw_row


class PyexcelExportEncoder(json.JSONEncoder):
    def __init__(self, *args, **kwargs):
        super(PyexcelExportEncoder, self).__init__(*args, **kwargs)
        self.kwargs = dict(kwargs)
        del self.kwargs['indent']
        self._replacement_map = {}

    def default(self, o):
        if isinstance(o, RowExport):
            key = uuid.uuid4().hex
            self._replacement_map[key] = json.dumps(o.value, **self.kwargs)
            return "@@%s@@" % (key,)
        elif isinstance(o, BytesIO):
            return base64.b64encode(o.getvalue()).decode()
        else:
            return super(PyexcelExportEncoder, self).default(o)

    def encode(self, o):
        result = super(PyexcelExportEncoder, self).encode(o)
        for k, v in self._replacement_map.items():
            result = result.replace('"@@%s@@"' % (k,), v)
        return result


class MyEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, RowExport):
            return o.data
        elif isinstance(o, BytesIO):
            return base64.b64encode(o.getvalue()).decode()

        return json.JSONEncoder.default(self, o)
