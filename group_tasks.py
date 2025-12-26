from collections import defaultdict

def agrupar_por_pessoa(tasks):
    tasks_por_pessoa = defaultdict(list)

    for task in tasks:
        pessoa = task.get("Assigned", "Sem responsável")
        tasks_por_pessoa[pessoa].append(task)

    return tasks_por_pessoa