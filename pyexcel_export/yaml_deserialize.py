import yaml

from collections import OrderedDict


class MyYamlLoader(yaml.SafeLoader):
    def construct_python_odict(self, node):
        return OrderedDict(self.construct_sequence(node))


MyYamlLoader.add_constructor(
    u'tag:yaml.org,2002:python/object/apply:collections.OrderedDict',
    MyYamlLoader.construct_python_odict)
