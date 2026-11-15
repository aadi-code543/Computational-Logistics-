import math
import re
import sys
from itertools import product
from pathlib import Path

import openpyxl


def rows(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    result = []
    for r in ws.iter_rows(values_only=True):
        result.append(list(r))
    return result


def load(d):
    f = lambda name: Path(d) / f"supply chain - data - {name}.xlsx"
    D = {}

    #materials required
    r = rows(f("material requirements"))
    D["mats"] = r[0][1:]               
    D["prods"] = []                      
    for x in r[1:]:
        D["prods"].append(x[0])
    D["req"] = {}
    for x in r[1:]:
        for i, m in enumerate(D["mats"]):
            D["req"][x[0], m] = x[1 + i] or 0

    #suppliers
    r = rows(f("suppliers"))
    D["sups"] = []
    for h in r[0][2::2]:
        D["sups"].append(re.match(r"Cost \((.*)\)", h).group(1))
    D["cost"], D["max"] = {}, {}
    for x in r[1:]:
        for i, s in enumerate(D["sups"]):
            D["cost"][s, x[0]] = x[2 + 2 * i]
            D["max"][s, x[0]] = x[3 + 2 * i]

    #processors
    r = rows(f("processors"))
    n = len(D["prods"])
    D["procs"] = []
    D["time"], D["rate"], D["cap"] = {}, {}, {}
    for x in r[1:]:
        D["procs"].append(x[0])
        for i, k in enumerate(D["prods"]):
            D["time"][x[0], k] = x[1 + i]
        D["rate"][x[0]] = x[1 + n]
        D["cap"][x[0]] = x[2 + n]

    #distributors
    r = rows(f("distributors"))
    D["dists"] = []                     
    D["demand"], D["price"] = {}, {}
    for x in r[1:]:
        D["dists"].append(x[0])
        for i, k in enumerate(D["prods"]):
            D["demand"][x[0], k] = x[1 + 2 * i]
            D["price"][x[0], k] = x[2 + 2 * i]

    #distance
    r = rows(f("distances"))
    D["dist"] = {}
    for x in r[1:]:                         
        for h, v in zip(r[0][1:], x[1:]):    
            if v is not None:
                D["dist"][x[0], h] = v
                D["dist"][h, x[0]] = v

    #shipping
    D["ship"] = {}                   
    for x in rows(f("shipping costs"))[1:]:
        D["ship"][x[0]] = x[2]
    return D



def parse_label(kind, label):
    left, item = label.split("\\")
    if kind == "supply":                   
        return "sup", (left, item)
    if kind == "process":
        return "proc", (left, item)
    a, b = left.split("→")
    if kind.startswith("ship (supply"):
        return "s2p", (a, b, item)
    return "p2d", (a, b, item)


def read_solution(path):
    sol = {"sup": {}, "s2p": {}, "proc": {}, "p2d": {}}
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    for kind, label, q in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        if kind:                            
            name, key = parse_label(kind, label)
            sol[name][key] = q or 0            
    return sol


def write_solution(template, out, sol):
    wb = openpyxl.load_workbook(template)       
    for kind, label, cell in wb.worksheets[0].iter_rows(min_row=2, max_col=3):
        if kind.value:
            name, key = parse_label(kind.value, label.value)
            v = sol[name].get(key, 0)           
            if name in ("sup", "s2p"):
                cell.value = round(v, 1)
            else:
                cell.value = int(round(v))
    wb.save(out)                             



# d - data, sol - solution, print/no print
def profit(D, sol, feedback=False):
    # bought from supplier, supplier to producer, producer to distributor
    sup, s2p, proc, p2d = sol["sup"], sol["s2p"], sol["proc"], sol["p2d"]

    #adds all money

    #producer to distributor
    rev = 0
    for (p, d, k), q in p2d.items():
        rev += q * D["price"][d, k]

    #supplier
    mat = 0
    for (s, m), q in sup.items():
        mat += q * D["cost"][s, m]

    #supplier to processor
    sh1 = 0
    for (s, p, m), q in s2p.items():
        sh1 += q * D["dist"][s, p] * D["ship"][m]

    #processing
    work = 0
    for (p, k), q in proc.items():
        work += q * D["time"][p, k] * D["rate"][p]

    #processor to distributor
    sh2 = 0
    for (p, d, k), q in p2d.items():
        sh2 += q * D["dist"][p, d] * D["ship"][k]


    #print the calc
    if feedback:
        lines = [("revenue", rev), ("materials", -mat), ("shipping 1", -sh1),
                 ("processing", -work), ("shipping 2", -sh2)]
        for name, v in lines:
            print(f"  {name:<11} {v:12,.2f}")
    return rev - mat - sh1 - work - sh2

#checks for rule breaks
def check(D, sol, tol=1e-6):
    sup, s2p, proc, p2d = sol["sup"], sol["s2p"], sol["proc"], sol["p2d"]
    S, P, Dd, M, K = D["sups"], D["procs"], D["dists"], D["mats"], D["prods"]
    get = lambda dct, *key: dct.get(key, 0)      
    bad = []                                   

    #cant cross supplier limit
    for s, m in product(S, M):
        bought = get(sup, s, m)
        if bought > D["max"][s, m] + tol:
            bad.append(f"buys {bought} {m} from {s}, max is {D['max'][s, m]}")
        shipped = 0
        for p in P:
            shipped += get(s2p, s, p, m)
        if shipped > bought + tol:
            bad.append(f"ships more {m} from {s} than was bought")

    #cant ship more to processor than bought
    for p in P:
        for m in M:
            have = 0                        
            for s in S:
                have += get(s2p, s, p, m)
            need = 0                          
            for k in K:
                need += D["req"][k, m] * get(proc, p, k)
            if need > have + tol:
                bad.append(f"{p} needs {need:.1f} {m} but receives {have:.1f}")

        #max processing capacity
        hours = 0
        for k in K:
            hours += D["time"][p, k] * get(proc, p, k)
        if hours > D["cap"][p] + tol:
            bad.append(f"{p} works {hours:.2f} h, capacity is {D['cap'][p]}")


        #max ship capacity
        for k in K:
            shipped = 0
            for d in Dd:
                shipped += get(p2d, p, d, k)
            if shipped > get(proc, p, k) + tol:
                bad.append(f"{p} ships more {k} than it made")



    #demand capped
    for d, k in product(Dd, K):
        got = 0
        for p in P:
            got += get(p2d, p, d, k)
        if got > D["demand"][d, k] + tol:
            bad.append(f"{d} gets {got} {k}, demand is {D['demand'][d, k]}")
    return bad


# Solver
# simplex has a table
# intial rows are the constraints
# final is the sol
# keeps improving sol untill nothing possible
#           x    y   s1  s2 | rhs
#   row1    1    1   1   0  | 4  
#   row2    1    0   0   1  | 3     
#   z      -3   -2   0   0  | 0    
# Example
# z is the sol row
# if z has a negative number then raising it makes profit
# keep increasing a negative number untill not possible
# goes to next negative in sol if possible
# repeat untill nothing left


EPS = 1e-9
def pivot(T, basis, r, k):
    pivot_value = T[r][k]
    new_row = []
    for v in T[r]:
        new_row.append(v / pivot_value)
    T[r] = new_row
    for i in range(len(T)):
        row = T[i]
        if i != r and abs(row[k]) > EPS:
            f = row[k]
            new_row = []
            for a, b in zip(row, T[r]):
                new_row.append(a - f * b)
            T[i] = new_row
    basis[r] = k



def run_simplex(T, basis):
    while True:
        z = T[-1]
        k = None
        for j in range(len(z) - 1):
            if z[j] < -EPS:
                k = j
                break
        if k is None:
            return
        r = None
        best = None
        for i in range(len(T) - 1):
            if T[i][k] > EPS:
                candidate = (T[i][-1] / T[i][k], basis[i])
                if best is None or candidate < best:
                    r = i
                    best = candidate
        if r is None:
            raise RuntimeError("problem is unbounded")
        pivot(T, basis, r, k)


def simplex(c, A, b):
    m, n = len(A), len(c) 
    T = []
    for i in range(m):
        row = list(A[i])             
        for j in range(m):         
            if i == j:
                row.append(1.0)
            else:
                row.append(0.0)
        row.append(-1.0)         
        row.append(b[i])          
        T.append(row)

    basis = []    
    for i in range(m):
        basis.append(n + i)
    x0 = n + m                     

    # adds x0 to each constraint to check if possible to solve
    # starts at a big number and reduced to hopefully 0
    worst = 0                       
    for i in range(m):
        if b[i] < b[worst]:
            worst = i
    if b[worst] < 0:
        T.append([0.0] * x0 + [1.0, 0.0]) 
        pivot(T, basis, worst, x0)        
        run_simplex(T, basis)           
        if T[-1][-1] < -1e-6:
            return None  
        T.pop()                           
        if x0 in basis: 
            r = basis.index(x0)
            for j in range(x0):
                if abs(T[r][j]) > EPS:
                    pivot(T, basis, r, j)
                    break  
    for i in range(len(T)):
        T[i] = T[i][:x0] + T[i][x0 + 1:]

    #actual profit
    z = []
    for v in c:
        z.append(-v)
    z += [0.0] * (m + 1)

    for i, k in enumerate(basis):
        if z[k]:
            factor = z[k]
            new_z = []
            for a, t in zip(z, T[i]):
                new_z.append(a - factor * t)
            z = new_z
    T.append(z)
    run_simplex(T, basis)

    x = [0.0] * n
    for i, k in enumerate(basis):
        if k < n:
            x[k] = T[i][-1]
    return T[-1][-1], x






# helps "round" fractional answers
# upper and lower bound then squeexed
# split up into two
def branch_and_bound(c, A, b, int_vars):
    n = len(c)
    best_val, best_x = -math.inf, None 
    todo = [([], [])]  
    while todo:
        extra_A, extra_b = todo.pop()  
        res = simplex(c, A + extra_A, b + extra_b)
        if res is None or res[0] <= best_val + 1e-6:
            continue
        val, x = res
        j = None
        for v in int_vars:
            if abs(x[v] - round(x[v])) > 1e-6:
                j = v
                break
        if j is None:
            best_val, best_x = val, x
            continue
        f = math.floor(x[j])
        down = [0.0] * n
        down[j] = 1.0                
        up = [0.0] * n
        up[j] = -1.0                 
        todo.append((extra_A + [up], extra_b + [-(f + 1)]))
        todo.append((extra_A + [down], extra_b + [f]))
    if best_x is None:
        raise RuntimeError("no feasible solution")
    return best_val, best_x





#solving
def solve(D):
    #suppliers, processors, distributors, materials, products (K)
    S, P, Dd, M, K = D["sups"], D["procs"], D["dists"], D["mats"], D["prods"]
    cols = []
    for s, p, m in product(S, P, M):
        cols.append(("s2p", s, p, m))
    for p, k in product(P, K):
        cols.append(("proc", p, k))
    for p, d, k in product(P, Dd, K):
        cols.append(("p2d", p, d, k))
    col = {}
    for j, v in enumerate(cols):
        col[v] = j
    n = len(cols)
    c = [0.0] * n
    for s, p, m in product(S, P, M): 
        c[col["s2p", s, p, m]] = -(D["cost"][s, m] + D["dist"][s, p] * D["ship"][m])
    for p, k in product(P, K):
        c[col["proc", p, k]] = -D["time"][p, k] * D["rate"][p]
    for p, d, k in product(P, Dd, K): 
        c[col["p2d", p, d, k]] = D["price"][d, k] - D["dist"][p, d] * D["ship"][k]

    A, b = [], [] 




    def limit(terms, rhs): 
        row = [0.0] * n    
        for key, coef in terms:       
            row[col[key]] += coef
        A.append(row)
        b.append(rhs)
    for s, m in product(S, M):
        terms = []
        for p in P:
            terms.append((("s2p", s, p, m), 1))
        limit(terms, D["max"][s, m])
    for p, m in product(P, M):
        terms = []
        for k in K:
            terms.append((("proc", p, k), D["req"][k, m])) 
        for s in S:
            terms.append((("s2p", s, p, m), -1))        
        limit(terms, 0)

    for p in P:
        terms = []
        for k in K:
            terms.append((("proc", p, k), D["time"][p, k]))
        limit(terms, D["cap"][p])
    for p, k in product(P, K):
        terms = []
        for d in Dd:
            terms.append((("p2d", p, d, k), 1))
        terms.append((("proc", p, k), -1))
        limit(terms, 0)
    for d, k in product(Dd, K):
        terms = []
        for p in P:
            terms.append((("p2d", p, d, k), 1))
        limit(terms, D["demand"][d, k])
    whole = []
    for j, v in enumerate(cols):
        if v[0] != "s2p":
            whole.append(j)
    _, x = branch_and_bound(c, A, b, whole)




    sol = {"sup": {}, "s2p": {}, "proc": {}, "p2d": {}}
    for j, v in enumerate(cols):
        name = v[0]    
        key = v[1:]         
        if name == "s2p":
            sol[name][key] = round(x[j], 1)
        else:
            sol[name][key] = int(round(x[j]))
    for s, m in product(S, M):
        total = 0
        for p in P:
            total += sol["s2p"][s, p, m]
        sol["sup"][s, m] = round(total, 1)
    return sol


def print_plan(sol):
    sections = [("BUY", "sup", "{1} from {0}", "kg"),
                ("SHIP MATERIALS", "s2p", "{2}: {0} -> {1}", "kg"),
                ("MAKE", "proc", "{1} at {0}", "units"),
                ("SHIP PRODUCTS", "p2d", "{2}: {0} -> {1}", "units")]
    for title, name, fmt, unit in sections:
        print(title)
        for key, q in sol[name].items():
            if q:
                print(f"  {q:>8g} {unit:<5} {fmt.format(*key)}")


def main():
    if len(sys.argv) < 4 or sys.argv[1] not in ("solve", "eval"):
        print(__doc__)
        sys.exit(1)

    mode, data, f = sys.argv[1:4]
    D = load(data)
    if mode == "solve":
        sol = solve(D)
        write_solution(Path(data) / "solution - template.xlsx", f, sol)
        print_plan(sol)
        print(f"\nwrote {f}")
    sol = read_solution(f)
    bad = check(D, sol)
    print("constraint violations:", bad or "none")
    total = profit(D, sol, feedback=True)
    print(f"PROFIT: {total:,.2f}")
if __name__ == "__main__":
    main()
