import pstats
p = pstats.Stats('stats')
p.sort_stats("cumtime").print_stats()