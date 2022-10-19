def le_se_al_python(malparido_tomas):
    si_se_python = False
    for x in malparido_tomas:
        j = malparido_tomas.count(x)
        if j > 1:
            si_se_python = True
    print(si_se_python)
le_se_al_python([1, 6, 8, 1, 22, 33, 2])