def get_manual_entries():
    return [
        {
            "Cidade": "Cidade Exemplo",
            "Tipo": "Setor Interno",
            "Setor": "Nome do Setor",
            "Sigla": "SIGLA",
            "Titular": "Fulano de Tal",
            "Unidade (Prédio)": "Prédio Principal",
            "Telefone": "3333-3333",
            "URL": ""
        }
    ]

def set_city_into_unidade(entries):
    """
    Recebe uma lista de dicts e atualiza a chave 'Unidade (Prédio)' para 'Cidade - Prédio'
    """
    for reg in entries:
        reg['Unidade (Prédio)'] = f"{reg['Cidade']} - {reg['Unidade (Prédio)']}"
    return entries

def get_ip_ranges_mapping():
    """
    Retorna o mapeamento de faixas de IP (CIDR) para a localidade física formatada.
    """
    return {
        "10.111.10.0/24": "Cidade Exemplo",
        "192.168.1.0/24": "Sede Administrativa"
    }

def get_location_by_ip(ip: str) -> str:
    """
    Busca a localidade física com base no IP.
    """
    if not ip or ip in ["Acesso Negado", "Timeout", "Erro", "Não encontrado"]:
        return ""
        
    import ipaddress
    try:
        user_ip = ipaddress.ip_address(ip)
        mapping = get_ip_ranges_mapping()
        
        for network_str, location in mapping.items():
            if user_ip in ipaddress.ip_network(network_str):
                return location
                
        return ""
    except:
        return ""