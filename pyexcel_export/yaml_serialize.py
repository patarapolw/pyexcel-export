import yaml
from io import BytesIO
import base64


class PyExcelYamlLoader(yaml.SafeLoader):
    def __init__(self, s):
        super().__init__(s)

        self.add_constructor(u'tag:yaml.org,2002:python/object/new:_io.BytesIO', self.construct_bytes_io)

    @staticmethod
    def construct_bytes_io(node, node_value):
        return BytesIO(base64.b64decode(node_value.value[0][1].value[0].value))
