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

    outfile = gzip.GzipFile(fname,'w')

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
            outfile.write(cell.value.encode('utf-8') + b"\n")
        else:
            outfile.write((str(cell.value) + u"\n").encode('utf-8'))
        outfile.write(b"====" + b"\n")

    outfile.write(b"-----" + b"\n")

    for cell in range_cells:
        parse_cell_info(cell)

        if cell.range.is_pointer:
            outfile.write((json.dumps(cell.range.reference) + u"\n").encode('utf-8'))
        else:
            outfile.write((cell.range.name + u"\n").encode('utf-8'))

        outfile.write(b"====" + b"\n")
        outfile.write(b"====" + b"\n")

    # writing the edges
    outfile.write(b"edges" + b"\n")
    for source, target in self.G.edges():
        outfile.write((source.address() + SEP + target.address() + u"\n").encode('utf-8'))

    # writing the rest
    if self.outputs is not None:
        outfile.write(b"outputs" + b"\n")
        outfile.write((SEP.join(self.outputs) + u"\n").encode('utf-8'))
    if self.inputs is not None:
        outfile.write(b"inputs" + b"\n")
        outfile.write((SEP.join(self.inputs) + u"\n").encode('utf-8'))
    outfile.write(b"named_ranges" + b"\n")
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
    infile = gzip.GzipFile(fname, 'rb')
        
    for line in infile.read().splitlines():

        line= line.decode("utf-8")

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

    G = DiGraph()
    G.add_nodes_from(nodes)
    G.add_edges_from(edges)

    print("Graph loading done, %s nodes, %s edges, %s cellmap entries" % (len(G.nodes()),len(G.edges()),len(cellmap)))

    return (G, cellmap, named_ranges, pointers, outputs, inputs)

########### based on json #################
def dump_json(self, fname):
    data = self.asdict()

    with gzip.GzipFile(fname, 'w') as outfile:
        outfile.write(json.dumps(data).encode('utf-8'))


def load_json(fname):

    def _decode_list(data):
        rv = []
        for item in data:
            if isinstance(item, unicode) and unicode != str:
                item = item.encode('utf-8')
            elif isinstance(item, list) and unicode != str:
                item = _decode_list(item)
            elif isinstance(item, dict):
                item = _decode_dict(item)
            rv.append(item)
        return rv

    def _decode_dict(data):
        rv = {}
        for key, value in data.items():
            if isinstance(key, unicode) and unicode != str:
                key = key.encode('utf-8')
            if isinstance(value, unicode) and unicode != str:
                value = value.encode('utf-8')
            elif isinstance(value, list):
                value = _decode_list(value)
            elif isinstance(value, dict):
                value = _decode_dict(value)
            rv[key] = value
        return rv

    with gzip.GzipFile(fname, 'r') as infile:
        data = json.loads(infile.read().decode('utf-8'), object_hook=_decode_dict)

    return data


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
