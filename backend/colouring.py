def greedy_coloring(graph: dict) -> dict:
    """
    Greedy graph coloring sorted by degree (descending).
    Assigns the smallest available slot (color) to each course.
    
    Returns:
        result (dict): {course: slot_number}
    """
    result = {}

    # Sort nodes by degree descending (most conflicted courses scheduled first)
    nodes = sorted(graph, key=lambda x: len(graph[x]), reverse=True)

    for node in nodes:
        # Collect colors already used by neighbors
        used_colors = {
            result[neighbor]
            for neighbor in graph[node]
            if neighbor in result
        }

        # Assign the smallest non-conflicting color (slot)
        color = 0
        while color in used_colors:
            color += 1

        result[node] = color

    return result


def dsatur_coloring(graph: dict) -> dict:
    """
    DSATUR algorithm — improved greedy coloring.
    Selects the node with the highest 'saturation degree' 
    (number of differently colored neighbors) at each step.
    Generally produces fewer slots than basic greedy.

    Returns:
        result (dict): {course: slot_number}
    """
    result = {}
    saturation = {node: 0 for node in graph}
    degree = {node: len(graph[node]) for node in graph}
    uncolored = set(graph.keys())

    while uncolored:
        # Pick node with highest saturation; break ties by degree
        node = max(
            uncolored,
            key=lambda x: (saturation[x], degree[x])
        )

        used_colors = {
            result[neighbor]
            for neighbor in graph[node]
            if neighbor in result
        }

        color = 0
        while color in used_colors:
            color += 1

        result[node] = color
        uncolored.remove(node)

        # Update saturation of uncolored neighbors
        for neighbor in graph[node]:
            if neighbor in uncolored:
                neighbor_colors = {
                    result[n]
                    for n in graph[neighbor]
                    if n in result
                }
                saturation[neighbor] = len(neighbor_colors)

    return result


def get_slot_summary(slots: dict) -> dict:
    """Group courses by slot for summary output."""
    summary = {}
    for course, slot in slots.items():
        summary.setdefault(slot, []).append(course)
    return summary
