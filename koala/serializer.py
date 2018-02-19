from __future__ import absolute_import, print_function

import json
import gzip
import networkx

from networkx.classes.digraph import DiGraph
from networkx.readwrite import json_graph
from networkx.drawing.nx_pydot import write_dot
from openpyxl.compat import unicode

from koala.Cell import Cell
from koala.Range import RangeCore, RangeFactory

SEP = ";;"

########### based on custom format #################
def dump(self, fname):
    outfile = gzip.GzipFile(fname, 'w')


    # write simple cells first
    simple_cells = [cell for cell in list(self.cellmap.values()) if cell.is_range == False]
    range_cells = [cell for cell in list(self.cellmap.values()) if cell.is_range]
    compiled_expressions = {}

    def parse_cell_info(cell):
        formula = cell.formula if cell.formula else "0"
        python_expression = cell.python_expression if cell.python_expression else "0"
        should_eval = cell.should_eval
        is_range = "1" if cell.is_range else "0"
        is_named_range = "1" if cell.is_named_range else "0"
        if cell.is_range:
            is_pointer = "1" if cell.range.is_pointer else "0"
        else:
            is_pointer = "0"

        compiled_expressions[cell.address()] = cell.compiled_expression

        # write common attributes
        outfile.write((SEP.join([
            cell.address(),
            formula,
            python_expression,
            is_range,
            is_named_range,
            is_pointer,
            should_eval
        ]) + u"\n").encode('utf-8'))

    for cell in simple_cells:
        parse_cell_info(cell)

        value = cell.value
        if isinstance(value, unicode):
            outfile.write(cell.value.encode('utf-8') + "\n")
        else:
            outfile.write(str(cell.value) + "\n")
        outfile.write("====" + "\n")

    outfile.write("-----" + "\n")

    for cell in range_cells:
        parse_cell_info(cell)

        if cell.range.is_pointer:
            outfile.write((json.dumps(cell.range.reference) + u"\n").encode('utf-8'))
        else:
            outfile.write((cell.range.name + u"\n").encode('utf-8'))

        outfile.write("====" + "\n")
        outfile.write("====" + "\n")

    # writing the edges
    outfile.write("edges" + "\n")
    for source, target in self.G.edges():
        outfile.write((source.address() + SEP + target.address() + u"\n").encode('utf-8'))

    # writing the rest
    if self.outputs is not None:
        outfile.write("outputs" + "\n")
        outfile.write((SEP.join(self.outputs) + u"\n").encode('utf-8'))
    if self.inputs is not None:
        outfile.write("inputs" + "\n")
        outfile.write((SEP.join(self.inputs) + u"\n").encode('utf-8'))
    outfile.write("named_ranges" + "\n")
    for k in self.named_ranges:
        outfile.write((k + SEP + self.named_ranges[k] + u"\n").encode('utf-8'))


    outfile.close()

def load(fname):

    def clean_bool(string):
        if string == "0":
            return None
        else:
            return string

    def to_bool(string):
        if string == "1" or string == "True":
            return True
        elif string == "0" or string == "False":
            return False
        else:
            return string
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
    pointers = set()
    outputs = None
    inputs = None
    named_ranges = {}
    infile = gzip.GzipFile(fname, 'r')

    for line in infile.read().splitlines():

        if line == "====":
            mode = "node0"
            continue
        if line == "-----":
            cellmap_temp = {n.address(): n for n in nodes}
            Range = RangeFactory(cellmap_temp)
            mode = "node0"
            continue
        elif line == "edges":
            cellmap = {n.address(): n for n in nodes}
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
            [address, formula, python_expression, is_range, is_named_range, is_pointer, should_eval] = line.split(SEP)
            formula = clean_bool(formula)
            python_expression = clean_bool(python_expression)
            is_range = to_bool(is_range)
            is_named_range = to_bool(is_named_range)
            is_pointer = to_bool(is_pointer)
            should_eval = should_eval
            mode = "node1"
        elif mode == "node1":
            if is_range:

                reference = json.loads(line) if is_pointer else line # in order to be able to parse dicts
                vv = Range(reference)

                if is_pointer:
                    if not is_named_range:
                        address = vv.name

                    pointers.add(address)

                cell = Cell(address, None, vv, formula, is_range, is_named_range, should_eval)
                cell.python_expression = python_expression
                nodes.append(cell)
            else:
                value = to_bool(to_float(line))

                cell = Cell(address, None, value, formula, is_range, is_named_range, should_eval)

                cell.python_expression = python_expression
                if formula:
                    if 'OFFSET' in formula or 'INDEX' in formula:
                        pointers.add(address)


                    cell.compile()
                nodes.append(cell)
        elif mode == "edges":
            source, target = line.split(SEP)
            edges.append((cellmap[source], cellmap[target]))
        elif mode == "outputs":
            outputs = line.split(SEP)
        elif mode == "inputs":
            inputs = line.split(SEP)
        elif mode == "named_ranges":
            k,v = line.split(SEP)
            named_ranges[k] = v

    G = DiGraph(data = edges)

    print("Graph loading done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap)))

    return (G, cellmap, named_ranges, pointers, outputs, inputs)

########### based on json #################
def dump_json(self, fname):
    data = json_graph.node_link_data(self.G)
    # save nodes as simple objects
    nodes = []
    for node in data["nodes"]:
        cell = node["id"]

        if isinstance(cell.range, RangeCore):
            range = cell.range
            value = {
                "cells": range.addresses,
                "values": range.values,
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
            "should_eval": cell.should_eval
        }]
    data["nodes"] = nodes
    data["outputs"] = self.outputs
    data["inputs"] = self.inputs
    data["named_ranges"] = self.named_ranges
    with gzip.GzipFile(fname, 'w') as outfile:
        outfile.write(json.dumps(data))


def load_json(fname):

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
        for key, value in data.items():
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
        new_cell = Cell(d["address"], None, value=value, formula=d["formula"], is_range = cell_is_range, is_named_range=d["is_named_range"], should_eval=d["should_eval"])
        new_cell.python_expression = d["python_expression"]
        new_cell.compile()
        return {"id": new_cell}

    nodes = list(map(cell_from_dict, data["nodes"]))
    data["nodes"] = nodes

    G = json_graph.node_link_graph(data)
    cellmap = {n.address():n for n in G.nodes()}

    print("Graph loading done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap)))

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
