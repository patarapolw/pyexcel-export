from datetime import datetime
from collections import OrderedDict


DEFAULT_META = OrderedDict([
    ('created', datetime.fromtimestamp(datetime.now().timestamp()).isoformat())
])
