# MÓDULO PARA TRATAMENTO DE CPF


def tratar_cpf(cpf):

    cpf = str(cpf)

    if len(cpf) == 9:

        cpf_tratado = f"00{cpf[0:1]}.{cpf[1:4]}.{cpf[4:7]}-{cpf[7:]}"

    elif len(cpf) == 10:

        cpf_tratado = f"0{cpf[0:2]}.{cpf[2:5]}.{cpf[5:8]}-{cpf[8:]}"

    elif len(cpf) == 11:

        cpf_tratado = f"{cpf[0:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

    else:
        cpf_tratado = cpf
    
    return cpf_tratado
