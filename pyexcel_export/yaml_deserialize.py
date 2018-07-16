import yaml
from io import BytesIO


class PyExcelYamlLoader(yaml.SafeLoader):
    def construct_bytes_io(self, node):
        return tuple(self.construct_sequence(node))

#
# PyExcelYamlLoader.add_constructor(
#     u'tag:yaml.org,2002:python/tuple',
#     PyExcelYamlLoader.construct_python_tuple)
