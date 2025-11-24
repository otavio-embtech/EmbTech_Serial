import struct
import crcmod

# Inicializa a função de cálculo CRC Modbus
_crc16 = crcmod.mkCrcFun(0x18005, rev=True, initCrc=0xFFFF, xorOut=0x0000)

def calculate_crc16(data: bytes) -> bytes:
    """
    Calcula o CRC16 para um quadro Modbus RTU.
    Retorna o CRC como dois bytes (low byte primeiro, high byte depois).
    """
    crc = _crc16(data)
    return crc.to_bytes(2, 'little') # 'little' para low byte primeiro

def hex_string_to_int(hex_str: str) -> int:
    """Converte uma string hexadecimal para um inteiro."""
    return int(hex_str, 16)

def bin_string_to_int(bin_str: str) -> int:
    """Converte uma string binária para um inteiro."""
    return int(bin_str, 2)

def oct_string_to_int(oct_str: str) -> int:
    """Converte uma string octal para um inteiro."""
    return int(oct_str, 8)

def convert_value_to_bytes(value_str: str, value_format: str, value_type: str) -> bytes:
    """
    Converte um valor de string para bytes com base no formato e tipo especificados.
    Suporta HEX, BIN, DEC Signed/Unsigned, OCT, e tipos Coil, Register, Float, Double.
    """
    if not value_str:
        raise ValueError("Valor para escrever não pode ser vazio.")

    if value_type == "Coil (Boolean)":
        # Para Coil, o valor é 0xFF00 para ON e 0x0000 para OFF
        if value_format == "DEC Signed" or value_format == "DEC Unsigned":
            val = int(value_str)
        elif value_format == "HEX":
            val = hex_string_to_int(value_str)
        elif value_format == "BIN":
            val = bin_string_to_int(value_str)
        elif value_format == "OCT":
            val = oct_string_to_int(value_str)
        else:
            raise ValueError(f"Formato de valor '{value_format}' não suportado para Coil.")
        
        if val == 1:
            return b'\xFF\x00' # ON
        elif val == 0:
            return b'\x00\x00' # OFF
        else:
            raise ValueError("Valor para Coil deve ser 0 ou 1.")

    elif value_type == "Register (16-bit)":
        if value_format == "DEC Signed":
            val = int(value_str)
            return struct.pack('>h', val) # '>h' para short (2 bytes) assinado, big-endian
        elif value_format == "DEC Unsigned":
            val = int(value_str)
            return struct.pack('>H', val) # '>H' para unsigned short (2 bytes) não assinado, big-endian
        elif value_format == "HEX":
            val = hex_string_to_int(value_str)
            return val.to_bytes(2, 'big') # 2 bytes, big-endian
        elif value_format == "BIN":
            val = bin_string_to_int(value_str)
            return val.to_bytes(2, 'big')
        elif value_format == "OCT":
            val = oct_string_to_int(value_str)
            return val.to_bytes(2, 'big')
        else:
            raise ValueError(f"Formato de valor '{value_format}' não suportado para Register.")

    elif value_type == "Float (32-bit)":
        val = float(value_str)
        return struct.pack('>f', val) # '>f' para float (4 bytes), big-endian (2 registers)

    elif value_type == "Double (64-bit)":
        val = float(value_str)
        return struct.pack('>d', val) # '>d' para double (8 bytes), big-endian (4 registers)
    
    else:
        raise ValueError(f"Tipo de valor '{value_type}' não suportado.")

def convert_bytes_to_value(data_bytes: bytes, response_format: str, value_type: str):
    """
    Converte bytes de resposta Modbus para um valor legível com base no formato e tipo.
    """
    if not data_bytes:
        return ""

    if value_type == "Coil (Boolean)":
        if data_bytes == b'\xFF\x00':
            return "1 (ON)"
        elif data_bytes == b'\x00\x00':
            return "0 (OFF)"
        else:
            return f"Valor Coil inesperado: {data_bytes.hex()}"

    elif value_type == "Register (16-bit)":
        if len(data_bytes) < 2:
            raise ValueError("Dados insuficientes para um registro de 16 bits.")
        
        if response_format == "DEC Signed":
            return struct.unpack('>h', data_bytes[:2])[0]
        elif response_format == "DEC Unsigned":
            return struct.unpack('>H', data_bytes[:2])[0]
        elif response_format == "HEX":
            return data_bytes[:2].hex().upper()
        elif response_format == "BIN":
            return bin(int(data_bytes[:2].hex(), 16))[2:].zfill(16)
        elif response_format == "OCT":
            return oct(int(data_bytes[:2].hex(), 16))[2:]
        else:
            return data_bytes[:2].hex().upper() # Padrão para HEX

    elif value_type == "Float (32-bit)":
        if len(data_bytes) < 4:
            raise ValueError("Dados insuficientes para um float de 32 bits.")
        return struct.unpack('>f', data_bytes[:4])[0]

    elif value_type == "Double (64-bit)":
        if len(data_bytes) < 8:
            raise ValueError("Dados insuficientes para um double de 64 bits.")
        return struct.unpack('>d', data_bytes[:8])[0]
    
    else:
        return data_bytes.hex().upper() # Padrão para HEX se o tipo for desconhecido

def build_modbus_rtu_request(slave_id: int, function_code_hex: str, address: int, 
                              quantity_or_value_str: str = None, value_format: str = None, value_type: str = None) -> bytes:
    """
    Constrói um quadro de requisição Modbus RTU.
    slave_id: ID do escravo (1-247)
    function_code_hex: Código da função Modbus em string hexadecimal (ex: "03", "06")
    address: Endereço inicial (registro/coil)
    quantity_or_value_str: Quantidade de itens para leitura ou valor para escrita
    value_format: Formato do valor (HEX, BIN, DEC Signed, DEC Unsigned, OCT) para escrita
    value_type: Tipo do valor (Coil, Register, Float, Double) para escrita
    """
    if not (1 <= slave_id <= 247):
        raise ValueError("ID do escravo Modbus deve estar entre 1 e 247.")

    try:
        fc = int(function_code_hex, 16)
    except ValueError:
        raise ValueError(f"Código de função Modbus inválido: {function_code_hex}")

    # Cabeçalho básico: ID do escravo + Código da função
    adu = bytearray([slave_id, fc])
    
    # Endereço inicial (2 bytes)
    adu.extend(address.to_bytes(2, 'big')) # 'big' para big-endian

    # Lógica para diferentes códigos de função
    if fc in [0x01, 0x02, 0x03, 0x04]: # Read Coils, Read Discrete Inputs, Read Holding Registers, Read Input Registers
        if quantity_or_value_str is None:
            raise ValueError("Quantidade de itens é necessária para funções de leitura.")
        try:
            quantity = int(quantity_or_value_str)
            if not (1 <= quantity <= 2000): # Limites típicos para quantidade
                raise ValueError("Quantidade de itens para leitura fora da faixa válida (1-2000).")
            adu.extend(quantity.to_bytes(2, 'big'))
        except ValueError:
            raise ValueError("Quantidade de itens deve ser um número inteiro válido.")

    elif fc == 0x05: # Write Single Coil
        if quantity_or_value_str is None or value_format is None or value_type != "Coil (Boolean)":
            raise ValueError("Valor e tipo de valor 'Coil (Boolean)' são necessários para escrever um único Coil.")
        
        coil_value_bytes = convert_value_to_bytes(quantity_or_value_str, value_format, value_type)
        adu.extend(coil_value_bytes) # Deve ser b'\xFF\x00' ou b'\x00\x00'

    elif fc == 0x06: # Write Single Register
        if quantity_or_value_str is None or value_format is None or value_type != "Register (16-bit)":
            raise ValueError("Valor e tipo de valor 'Register (16-bit)' são necessários para escrever um único Register.")
        
        register_value_bytes = convert_value_to_bytes(quantity_or_value_str, value_format, value_type)
        adu.extend(register_value_bytes) # Deve ser 2 bytes

    else:
        raise ValueError(f"Código de função Modbus {function_code_hex} não suportado para construção de requisição.")

    # Adiciona CRC
    crc = calculate_crc16(adu)
    adu.extend(crc)

    return bytes(adu)

def parse_modbus_rtu_response(response_bytes: bytes, function_code_hex: str, 
                              expected_quantity: int = None, value_type: str = None, response_format: str = None) -> (bool, str, any):
    print(f"DEBUG MODBUS LIB: Resposta BRUTA da serial: {response_bytes.hex().upper()} (Tamanho: {len(response_bytes)} bytes)")
    """
    Analisa um quadro de resposta Modbus RTU.
    Retorna (True/False para sucesso/falha, mensagem de erro, valor extraído).
    """
    if not response_bytes or len(response_bytes) < 5: # Mínimo: Slave ID, FC, Byte Count (ou Valor), CRC (2 bytes)
        return False, "Resposta Modbus muito curta ou vazia.", None

    # Verifica CRC
    received_crc = response_bytes[-2:]
    calculated_crc = calculate_crc16(response_bytes[:-2])
    if received_crc != calculated_crc:
        return False, f"Erro de CRC. Recebido: {received_crc.hex().upper()}, Calculado: {calculated_crc.hex().upper()}", None

    slave_id = response_bytes[0]
    fc = response_bytes[1]
    
    try:
        expected_fc = int(function_code_hex, 16)
    except ValueError:
        return False, f"Código de função esperado inválido: {function_code_hex}", None

    # Verifica se é uma resposta de erro (código de função + 0x80)
    if fc == (expected_fc + 0x80):
        exception_code = response_bytes[2]
        exception_messages = {
            0x01: "Illegal Function (Função Ilegal)",
            0x02: "Illegal Data Address (Endereço de Dados Ilegal)",
            0x03: "Illegal Data Value (Valor de Dados Ilegal)",
            0x04: "Slave Device Failure (Falha no Dispositivo Escravo)",
            0x05: "Acknowledge (Reconhecimento)",
            0x06: "Slave Device Busy (Dispositivo Escravo Ocupado)",
            0x08: "Memory Parity Error (Erro de Paridade de Memória)",
            0x0A: "Gateway Path Unavailable (Caminho do Gateway Indisponível)",
            0x0B: "Gateway Target Device Failed to Respond (Dispositivo Alvo do Gateway Não Respondeu)"
        }
        error_msg = exception_messages.get(exception_code, f"Exceção Modbus desconhecida: 0x{exception_code:02X}")
        return False, f"Resposta de Exceção Modbus: {error_msg}", None

    # Verifica se o código de função corresponde ao esperado
    if fc != expected_fc:
        return False, f"Código de função na resposta ({fc:02X}) não corresponde ao esperado ({expected_fc:02X}).", None

    # Extrai dados com base no código de função
    if fc in [0x01, 0x02, 0x03, 0x04]: # Read Coils, Read Discrete Inputs, Read Holding Registers, Read Input Registers
        if len(response_bytes) < 3: # Slave ID, FC, Byte Count
            return False, "Resposta de leitura muito curta.", None
        
        byte_count = response_bytes[2]
        data_bytes = response_bytes[3:-2] # Dados + CRC
        
        if len(data_bytes) != byte_count:
            return False, f"Contagem de bytes ({byte_count}) não corresponde ao comprimento dos dados ({len(data_bytes)}) na resposta de leitura.", None
        
        # Converte os bytes de dados para o tipo de valor esperado
        if value_type == "Coil (Boolean)":
            # Coils são empacotados em bits, 8 coils por byte.
            # O expected_quantity é o número de coils esperado.
            extracted_values = []
            for i in range(expected_quantity):
                byte_index = i // 8
                bit_index = i % 8
                if byte_index < len(data_bytes):
                    is_set = (data_bytes[byte_index] >> bit_index) & 0x01
                    extracted_values.append(1 if is_set else 0)
                else:
                    return False, f"Dados insuficientes para o número de coils esperado ({expected_quantity}).", None
            return True, "Resposta de leitura de coils analisada com sucesso.", extracted_values
        
        elif value_type == "Register (16-bit)":
            # Registros de 16 bits (2 bytes por registro)
            if expected_quantity is None:
                return False, "Quantidade esperada é necessária para parsing de Registros.", None
            
            expected_data_len = expected_quantity * 2
            if len(data_bytes) < expected_data_len:
                return False, f"Dados insuficientes para o número de registros esperado ({expected_quantity}).", None
            
            extracted_values = []
            for i in range(expected_quantity):
                reg_bytes = data_bytes[i*2 : (i*2)+2]
                try:
                    val = convert_bytes_to_value(reg_bytes, response_format, value_type)
                    extracted_values.append(val)
                except ValueError as e:
                    return False, f"Erro ao converter registro {i+1}: {e}", None
            return True, "Resposta de leitura de registros analisada com sucesso.", extracted_values

        elif value_type == "Float (32-bit)":
            # Float de 32 bits (4 bytes, 2 registros)
            if len(data_bytes) < 4:
                return False, "Dados insuficientes para um float de 32 bits.", None
            try:
                val = convert_bytes_to_value(data_bytes[:4], response_format, value_type)
                return True, "Resposta de leitura de float analisada com sucesso.", val
            except ValueError as e:
                return False, f"Erro ao converter float: {e}", None

        elif value_type == "Double (64-bit)":
            # Double de 64 bits (8 bytes, 4 registros)
            if len(data_bytes) < 8:
                return False, "Dados insuficientes para um double de 64 bits.", None
            try:
                val = convert_bytes_to_value(data_bytes[:8], response_format, value_type)
                return True, "Resposta de leitura de double analisada com sucesso.", val
            except ValueError as e:
                return False, f"Erro ao converter double: {e}", None
        else:
            return False, f"Tipo de valor '{value_type}' não suportado para parsing de leitura Modbus.", None

    elif fc in [0x05, 0x06]: # Write Single Coil, Write Single Register (resposta ecoa a requisição)
        # Respostas de escrita ecoam a requisição (Slave ID, FC, Address, Value, CRC)
        if len(response_bytes) < 6: # Slave ID, FC, Address (2), Value (2), CRC (2)
            return False, "Resposta de escrita muito curta.", None
        
        # O valor retornado é o valor escrito, que está no quadro de resposta.
        # Para 0x05, o valor é 0xFF00 ou 0x0000.
        # Para 0x06, o valor é o valor de 2 bytes escrito.
        written_address = int.from_bytes(response_bytes[2:4], 'big')
        written_value_bytes = response_bytes[4:-2]

        # Para fins de validação, podemos apenas verificar se a resposta é um eco válido.
        # Não precisamos converter o valor de volta, pois a validação já foi feita na requisição.
        return True, "Comando de escrita Modbus confirmado.", written_value_bytes.hex().upper()

    else:
        return False, f"Código de função Modbus {fc:02X} não suportado para análise de resposta.", None

