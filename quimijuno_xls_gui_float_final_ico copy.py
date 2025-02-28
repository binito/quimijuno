import os

def read_products_file(filename):
    """Reads product information from a file."""
    products = []
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            for line in lines:
                # Skip empty lines
                if not line.strip():
                    continue
                # Parse the line into product data
                reference, description, quantity_str, unit_type, price_str, taxes_str, amount_str = line.split(' ', maxsplit=6)
                
                # Try-except block for parsing numerical values
                try:
                    quantity = float(quantity_str)
                    price = float(price_str)
                    taxes = float(taxes_str) if taxes_str else 0.0
                    amount = float(amount_str)
                except ValueError:
                    print(f"Skipping invalid data in line: {line}")
                    continue
                
                products.append({
                    'REFERÊNCIA': reference,
                    'DESCRIÇÃO': description,
                    'QUANTIDADE': quantity,
                    'UNIDADE': unit_type,
                    'PREÇO UNITÁRIO': price,
                    'IMPOSTOS': taxes,
                    'AMOUNT': amount
                })
        
        return products
    
    except FileNotFoundError:
        print(f"Error: file {filename} not found.")
        return []
    except Exception as e:
        print(f"An error occurred while reading the file: {str(e)}")
        return []

def write_products_file(filename, products):
    """Saves product data to a formatted text file."""
    with open(filename, 'w', encoding='utf-8') as file:
        for product in products:
            # Ensure all values are converted to strings
            str_product = {
                'REFERÊNCIA': str(product.get('REFERÊNCIA', '')),
                'DESCRIÇÃO': str(product.get('DESCRIÇÃO', '')),
                'QUANTIDADE': float2str(product.get('QUANTIDADE', 0.0)),
                'UNIDADE': product.get('UNIDADE', ''),
                'PREÇO UNITÁRIO': float2str(product.get('PREÇO UNITÁRIO', 0.0)),
                'IMPOSTOS': float2str(product.get('IMPOSTOS', 0.0)),
                'AMOUNT': float2str(product.get('AMOUNT', 0.0))
            }
            # Format the string
            formatted_line = (
                f"{str_product['REFERÊNCIA']} {str_product['DESCRIÇÃO']} "
                f"{str_product['QUANTIDADE']} {str_product['UNIDADE']} "
                f"{str_product['PREÇO UNITÁRIO']} {str_product['IMPOSTOS']} "
                f"{str_product['AMOUNT']}"
            )
            file.write(formatted_line + '\n')

def float2str(value):
    return str(value).replace('.', '').replace('e', 'E') if '.' in str(value) else ''

def main():
    filename = 'produtos.txt'
    products = read_products_file(filename)
    print(f"Successfully read {len(products)} products.")
    
    # Modify the data structure to be more efficient
    flat_products = []
    for product in products:
        flat_product = {
            'REFERÊNCIA': product['REFERÊNCIA'],
            'DESCRIÇÃO': product['DESCRIÇÃO'],
            'QUANTIDADE': product['QUANTIDADE'],
            'UNIDADE': product['UNIDADE'],
            'PREÇO UNITÁRIO': product['PREÇO UNITÁRIO'],
            'IMPOSTOS': product['IMPOSTOS'],
            'AMOUNT': product['AMOUNT']
        }
        flat_products.append(flat_product)
    
    write_products_file(filename, flat_products)

if __name__ == "__main__":
    main()