import tkinter as tk
from tkinter import ttk
import pandas as pd
import networkx as nx
import plotly.graph_objs as go
from plotly.offline import plot

def build_graph_from_excel(df):
    G = nx.DiGraph()
    unique_groups = df["F"].unique()
    color_palette = [
        "#F94144", "#F3722C", "#F8961E", "#F9C74F", "#90BE6D",
        "#43AA8B", "#577590", "#277DA1", "#4D908E", "#F9844A"
    ]
    group_to_color = {}
    for i, grp in enumerate(unique_groups):
        group_to_color[grp] = color_palette[i % len(color_palette)]
    for _, row in df.iterrows():
        edge_name = str(row["A"])
        in_process = str(row["B"])
        in_owner = str(row["C"])
        out_process = str(row["D"])
        out_owner = str(row["E"])
        group_val = str(row["F"])
        if in_process not in G:
            G.add_node(in_process, owner=in_owner, group=group_val)
        if out_process not in G:
            G.add_node(out_process, owner=out_owner, group=group_val)
        G.add_edge(out_process, in_process, name=edge_name)
    for node in G.nodes():
        grp = G.nodes[node].get("group", "")
        node_color = group_to_color.get(grp, "#999999")
        G.nodes[node]["node_color"] = node_color
    return G, group_to_color

def build_subgraph_for_nodes(G, node_list):
    # Функция, которая объединяет "окрестности" нескольких процессов
    all_nodes = set()
    for center_node in node_list:
        if center_node not in G:
            continue
        all_nodes.add(center_node)
        for pred in G.predecessors(center_node):
            all_nodes.add(pred)
        for succ in G.successors(center_node):
            all_nodes.add(succ)
    subG = G.subgraph(all_nodes).copy()
    return subG

def create_plotly_figure(G, group_to_color, title="Plotly Graph"):
    pos = nx.spring_layout(G, k=0.7, seed=42)
    node_x = []
    node_y = []
    node_text = []
    node_color = []
    for node in G.nodes():
        x, y = pos[node]
        node_x.append(x)
        node_y.append(y)
        txt = f"{node}<br>Владелец: {G.nodes[node].get('owner','')}"
        node_text.append(txt)
        node_color.append(G.nodes[node]["node_color"])
    node_trace = go.Scatter(
        x=node_x,
        y=node_y,
        mode='markers',
        text=node_text,
        hoverinfo='text',
        marker=dict(
            showscale=False,
            color=node_color,
            size=12,
            line_width=1
        ),
        name="Узлы"
    )
    edge_x = []
    edge_y = []
    for u, v in G.edges():
        x0, y0 = pos[u]
        x1, y1 = pos[v]
        edge_x += [x0, x1, None]
        edge_y += [y0, y1, None]
    edge_trace = go.Scatter(
        x=edge_x,
        y=edge_y,
        line=dict(width=1, color='#888'),
        hoverinfo='none',
        mode='lines',
        name="Связи (линии)"
    )
    mid_x = []
    mid_y = []
    mid_text = []
    for u, v in G.edges():
        x0, y0 = pos[u]
        x1, y1 = pos[v]
        mx = (x0 + x1) / 2
        my = (y0 + y1) / 2
        edge_label = G[u][v].get("name", "")
        mid_x.append(mx)
        mid_y.append(my)
        mid_text.append(edge_label)
    midpoint_trace = go.Scatter(
        x=mid_x,
        y=mid_y,
        text=mid_text,
        mode='markers',
        hoverinfo='text',
        marker=dict(size=10, color='rgba(0,0,0,0)', line_width=0),
        name="Подписи связей"
    )

    # Фиктивные трейсы для легенды (каждое значение F)
    legend_traces = []
    for grp, clr in group_to_color.items():
        dummy_trace = go.Scatter(
            x=[None], y=[None],
            mode='markers',
            marker=dict(size=10, color=clr),
            legendgroup=str(grp),
            showlegend=True,
            name=f"Группа {grp}"
        )
        legend_traces.append(dummy_trace)
    
    fig = go.Figure(
        data=[edge_trace, node_trace, midpoint_trace] + legend_traces,
        layout=go.Layout(
            title=title,
            showlegend=True,
            hovermode='closest',
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
        )
    )

    for u, v in G.edges():
        x0, y0 = pos[u]
        x1, y1 = pos[v]
        fig.add_annotation(
            x=x1,
            y=y1,
            ax=x0,
            ay=y0,
            xref='x',
            yref='y',
            axref='x',
            ayref='y',
            showarrow=True,
            arrowhead=3,
            arrowsize=1,
            arrowwidth=1,
            arrowcolor='#FF0000',  # контрастный цвет стрелки
            opacity=1
        )

    fig.update_layout(dragmode='zoom', autosize=True)
    return fig

def main():
    df = pd.read_excel("data.xlsx")
    G, group_to_color = build_graph_from_excel(df)
    unique_b = sorted(set(df["B"].dropna()))

    root = tk.Tk()
    root.title("Выбор процессов")

    label = tk.Label(root, text="Выберите один или несколько процессов:")
    label.pack(padx=10, pady=5)

    frame_list = tk.Frame(root)
    frame_list.pack(padx=10, pady=5, fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame_list)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(frame_list, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set)
    for proc in unique_b:
        listbox.insert(tk.END, proc)
    listbox.pack(side=tk.LEFT, fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    def on_build():
        selected_indices = listbox.curselection()
        selected_procs = [unique_b[i] for i in selected_indices]
        if len(selected_procs) == 0:
            fig = create_plotly_figure(G, group_to_color, title="Полный граф")
            plot(fig, filename="plotly_network_all.html", auto_open=True)
        else:
            subG = build_subgraph_for_nodes(G, selected_procs)
            fig = create_plotly_figure(subG, group_to_color, title="Подграф для выбранных процессов")
            plot(fig, filename="plotly_subnetwork.html", auto_open=True)

    btn = tk.Button(root, text="Построить граф", command=on_build)
    btn.pack(padx=10, pady=5)

    label2 = tk.Label(root, text="Если не выбрано ничего - строится весь граф")
    label2.pack(padx=10, pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
