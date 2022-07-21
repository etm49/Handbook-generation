from enum import Enum

# statistics that will be calculated
class stats(Enum):
    minSids = 'SIDS minimum'
    minGlobal = 'Global minimum'
    maxSids = 'SIDS maximum'
    maxGlobal = 'Global maximum'
    aveSids = 'SIDS average'
    aveGlobal = 'Global average'
    indCount = 'Indicator Count for Dataset'

# values that will be inserted into the text
class outputs(Enum):
    countryName = 'name of country'
    value = 'statistical value calculated'
    year = 'year under consideration'