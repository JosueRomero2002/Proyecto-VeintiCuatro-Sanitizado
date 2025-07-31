entradas = [1, 2, 3]

pesos = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]

sesgos = [1, 2, 3]

salida = []

for i in range(0, len(entradas)):

    for j in range(0, len(pesos[0])):
        print(pesos[i][j])

for i in range(0, len(entradas)):
    arreglo = 0
    for j in range(0, len(pesos[0])):
        arreglo = arreglo + entradas[i] * pesos[i][j]
    arreglo = arreglo + sesgos[i]
    salida.append(arreglo)

print(salida)

salida = []


for n_sesgos, n_pesos in zip(sesgos, pesos):
    print([entrada * peso for (entrada, peso) in zip(entradas, n_pesos)])


for n_sesgos, n_pesos in zip(sesgos, pesos):
    salida.append(
        sum([entrada * peso for entrada, peso in zip(entradas, n_pesos)]) + n_sesgos
    )


print(salida)
