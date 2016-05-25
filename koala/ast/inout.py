import json
import gzip
import networkx

from koala.ast.excelutils import Cell
from Range import RangeCore

from networkx.classes.digraph import DiGraph
from networkx.readwrite import json_graph
from networkx.algorithms import number_connected_components
from networkx.drawing.nx_pydot import write_dot
from networkx.drawing.nx_pylab import draw, draw_circular

SEP = ";;"

########### based on custom format #################
def dump2(self, fname):
    data = json_graph.node_link_data(self.G)
    outfile = gzip.GzipFile(fname, 'w')
    for node in data["nodes"]:
        cell = node["id"]
        formula = cell.formula if cell.formula else "0"
        python_expression = cell.python_expression if cell.python_expression else "0"
        always_eval = "1" if cell.always_eval else "0"
        is_range = "1" if cell.is_range else "0"
        is_named_range = "1" if cell.is_named_range else "0"
        always_eval = "1" if cell.always_eval else "0"

        # write common attributes
        outfile.write(SEP.join([
            cell.address(),
            formula,
            python_expression,
            is_range,
            is_named_range,
            always_eval
        ]) + "\n")
        
        # write a bloc a just a value depending on is_range
        if isinstance(cell.range, RangeCore):
            range = cell.range
            outfile.write(SEP.join(map(str,range.value)) + "\n")
            outfile.write(SEP.join(map(str,range.cells)) + "\n")
            outfile.write(str(range.nrows) + SEP + str(range.ncols) + "\n")
        else:
            outfile.write(str(cell.value) + "\n")

        # end of the block corresponding to a cell
        outfile.write("====" + "\n")

    # writing the edges
    outfile.write("edges" + "\n")
    for edge in data["links"]:
        outfile.write(str(edge["source"]) + SEP + str(edge["target"]) + "\n")

    # writing the rest
    outfile.write("outputs" + "\n")
    outfile.write(SEP.join(self.outputs) + "\n")
    outfile.write("inputs" + "\n")
    outfile.write(SEP.join(self.inputs) + "\n")
    outfile.write("named_ranges" + "\n")
    for k in self.named_ranges:
        outfile.write(k + SEP + self.named_ranges[k] + "\n")
    
    outfile.close()

def load2(fname):

    def clean_bool(string):
        if string == "0":
            return None
        else:
            return string

    def to_bool(string):
        if string == "1":
            return True
        else:
            return False
    def to_float(string):
        if string == "None":
            return None
        try:
            return float(string)
        except:
            return string

    mode = "node0"
    nodes = []
    edges = []
    named_ranges = {}
    infile = gzip.GzipFile(fname, 'r')
    for line in infile.read().splitlines():

        if line == "====":
            mode = "node0"
            continue
        elif line == "edges":
            mode = "edges"
            continue
        elif line == "outputs":
            mode = "outputs"
            continue
        elif line == "inputs":
            mode = "inputs"
            continue
        elif line == "named_ranges":
            mode = "named_ranges"
            continue

        if mode == "node0":
            [address, formula, python_expression, is_range, is_named_range, always_eval] = line.split(SEP)
            formula = clean_bool(formula)
            python_expression = clean_bool(python_expression)
            is_range = to_bool(is_range)
            is_named_range = to_bool(is_named_range)
            always_eval = to_bool(always_eval)
            mode = "node1"
        elif mode == "node1":
            if is_range:
                splits = line.split(SEP)
                if len(splits) > 1:
                    values = map(to_float, splits)
                else:
                    try:
                        values = [to_float(line)]
                    except:
                        values = [None]
                mode = "node2"
            else:
                values = to_float(line)
                cell = Cell(address, None, values, formula, is_range, is_named_range, always_eval)
                cell.python_expression = python_expression
                cell.compile()                    
                nodes.append(cell)
        elif mode == "node2":
            splits = line.split(SEP)
            if len(splits) > 1:
                cells = splits
            else:
                cells = [address]
            mode = "node3"
        elif mode == "node3":
            nrows, ncols = map(int, line.split(SEP))

            if values == ['']:
                cells = []
                values = []
            vv = RangeCore(cells, values, nrows = nrows, ncols = ncols)
            cell = Cell(address, None, vv, formula, is_range, is_named_range, always_eval)
            cell.python_expression = python_expression
            cell.compile() 
            nodes.append(cell)

        elif mode == "edges":
            source, target = line.split(SEP)
            edges.append((int(source), int(target)))
        elif mode == "outputs":
            outputs = line.split(SEP)
        elif mode == "inputs":
            inputs = line.split(SEP)
        elif mode == "named_ranges":
            k,v = line.split(SEP)
            named_ranges[k] = v
            
    pos2cell = dict(enumerate(nodes))
    links = [(pos2cell[i], pos2cell[j]) for (i,j) in edges]
    print links[0]
    G = DiGraph(data = links)
    cellmap = {n.address():n for n in G.nodes()}

    print "Graph loading done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))

    return (G, cellmap, named_ranges, outputs, inputs)

########### based on json #################
def dump(self, fname):
    data = json_graph.node_link_data(self.G)
    # save nodes as simple objects
    nodes = []
    for node in data["nodes"]:
        cell = node["id"]

        if isinstance(cell.range, RangeCore):
            range = cell.range
            value = {
                "cells": range.cells,
                "values": range.value,
                "nrows": range.nrows,
                "ncols": range.ncols
            }
        else:
            value = cell.value

        nodes += [{
            "address": cell.address(),
            "formula": cell.formula,
            "value": value,
            "python_expression": cell.python_expression,
            "is_named_range": cell.is_named_range,
            "always_eval": cell.always_eval
        }]
    data["nodes"] = nodes
    data["outputs"] = self.outputs
    data["inputs"] = self.inputs
    data["named_ranges"] = self.named_ranges
    with gzip.GzipFile(fname, 'w') as outfile:
        outfile.write(json.dumps(data))


def load(fname):

    def _decode_list(data):
        rv = []
        for item in data:
            if isinstance(item, unicode):
                item = item.encode('utf-8')
            elif isinstance(item, list):
                item = _decode_list(item)
            elif isinstance(item, dict):
                item = _decode_dict(item)
            rv.append(item)
        return rv

    def _decode_dict(data):
        rv = {}
        for key, value in data.iteritems():
            if isinstance(key, unicode):
                key = key.encode('utf-8')
            if isinstance(value, unicode):
                value = value.encode('utf-8')
            elif isinstance(value, list):
                value = _decode_list(value)
            elif isinstance(value, dict):
                value = _decode_dict(value)
            rv[key] = value
        return rv
    with gzip.GzipFile(fname, 'r') as infile:
        data = json.loads(infile.read(), object_hook=_decode_dict)
    def cell_from_dict(d):
        cell_is_range = type(d["value"]) == dict
        if cell_is_range:
            range = d["value"]
            if len(range["values"]) == 0:
                range["values"] = [None] * len(range["cells"])
            value = RangeCore(range["cells"], range["values"], nrows = range["nrows"], ncols = range["ncols"])
        else:
            value = d["value"]
        new_cell = Cell(d["address"], None, value=value, formula=d["formula"], is_range = cell_is_range, is_named_range=d["is_named_range"], always_eval=d["always_eval"])
        new_cell.python_expression = d["python_expression"]
        new_cell.compile()
        return {"id": new_cell}

    nodes = map(cell_from_dict, data["nodes"])
    data["nodes"] = nodes

    G = json_graph.node_link_graph(data)
    cellmap = {n.address():n for n in G.nodes()}

    print "Graph loading done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap))

    return (G, cellmap, data["named_ranges"], data["outputs"], data["inputs"])


########### based on dot #################
def export_to_dot(self,fname):
    write_dot(self.G,fname)
                

########### plotting #################
def plot_graph(self):
    import matplotlib.pyplot as plt

    pos=networkx.spring_layout(self.G,iterations=2000)
    #pos=networkx.spectral_layout(G)
    #pos = networkx.random_layout(G)
    networkx.draw_networkx_nodes(self.G, pos)
    networkx.draw_networkx_edges(self.G, pos, arrows=True)
    networkx.draw_networkx_labels(self.G, pos)
    plt.show()