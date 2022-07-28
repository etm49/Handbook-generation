from enum import Enum

# statistics that will be calculated
class stats(Enum):
    minSids = 'SIDS minimum'
    minGlobal = 'Global minimum'
    maxSids = 'SIDS maximum'
    maxGlobal = 'Global maximum'
    aveSids = 'SIDS average'
    aveGlobal = 'Global average'
    minPacific = 'pacific minimum'
    maxPacific = 'pacific maximum'
    avePacific = 'pacific average'
    minCarribean = 'Carribean minimum'
    maxCarribean = 'Carribean maximum'
    aveCarribean = 'Carribean average'
    minAIS = 'AIS minimum'
    maxAIS = 'AIS maximum'
    aveAIS = 'AIS average'

# values that will be inserted into the text
class outputs(Enum):
    countryName = 'name of country'
    value = 'statistical value calculated'
    year = 'year under consideration'