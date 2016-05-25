import networkx

def subgraph(G, seed):
    subgraph = networkx.DiGraph()
    todo = map(lambda n: (seed,n), G.predecessors(seed))
    while len(todo) > 1:
        previous, current = todo.pop()
        addr = current.address()
        subgraph.add_node(current)
        subgraph.add_edge(previous, current)
        for n in G.predecessors(current):            
            if n not in subgraph.nodes():
                todo += [(current,n)]

    return subgraph