"""
graph.py — Standalone conflict graph builder.

NOTE: This module is NOT imported by app.py or the live pipeline.
The same build_conflict_graph() logic is inlined directly inside main.py
for the production pipeline (with an additional `all_courses` filter argument).

This file is kept as a self-contained utility for:
  - Unit testing the graph-building logic in isolation
  - Running get_graph_stats() on an existing graph without the full server

To test standalone:
    python -c "from graph import build_conflict_graph, get_graph_stats; g = build_conflict_graph({'S1': ['C1','C2']}); print(get_graph_stats(g))"
"""
from itertools import combinations


def build_conflict_graph(student_courses: dict) -> dict:
    """
    Build a conflict graph where:
    - Nodes = courses
    - Edges = two courses share at least one common student
    
    Returns:
        graph (dict): adjacency set representation
    """
    graph = {}

    # Initialize all courses as nodes
    for courses in student_courses.values():
        for course in courses:
            if course not in graph:
                graph[course] = set()

    # Add edges for every pair of courses a student takes
    for student_id, courses in student_courses.items():
        unique_courses = list(set(courses))  # deduplicate per student
        for c1, c2 in combinations(unique_courses, 2):
            graph[c1].add(c2)
            graph[c2].add(c1)

    return graph


def get_graph_stats(graph: dict) -> dict:
    """Return basic statistics about the conflict graph."""
    num_nodes = len(graph)
    num_edges = sum(len(neighbors) for neighbors in graph.values()) // 2
    max_degree = max((len(v) for v in graph.values()), default=0)
    avg_degree = (sum(len(v) for v in graph.values()) / num_nodes) if num_nodes else 0

    return {
        "courses": num_nodes,
        "conflict_edges": num_edges,
        "max_conflicts": max_degree,
        "avg_conflicts": round(avg_degree, 2),
    }
