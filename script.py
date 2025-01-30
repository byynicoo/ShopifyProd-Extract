import shopify
import pandas as pd

# Parametri APP
SHOPIFY_API_KEY = 'your shopify api key'
SHOPIFY_API_PASSWORD = 'ShopifyAPIpsw'
SHOPIFY_STORE_NAME = 'example.myshopify.com'

def connect_to_shopify():
    # Connection to shopify URL
    shop_url = f"https://{SHOPIFY_API_KEY}:{SHOPIFY_API_PASSWORD}@{SHOPIFY_STORE_NAME}"
    shopify.ShopifyResource.set_site(shop_url)

def fetch_products():
    # Fetch all products from Shopify store
    products = []
    page = 1

    while True:
        # Products research
        product_batch = shopify.Product.find(limit=250, page=page)
        if not product_batch:
            break
        products.extend(product_batch)
        page += 1

    return products

def extract_product_data(products):
    # Extract relevant product data
    product_data = []

    for product in products:
        product_info = {
            'Product ID': product.id,
            'Status': product.status,
            'Title': product.title,
            'Barcode': product.barcode,
            'Vendor': product.vendor,
            'Tags': product.tags,
        }
        product_data.append(product_info)

    return product_data

def save_to_excel(data, filename='shopify_products.xlsx'):
    # Save product data to an Excel file
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)

def main():
    # Main function to execute the workflow
    connect_to_shopify()
    products = fetch_products()
    product_data = extract_product_data(products)
    save_to_excel(product_data)
    print(f"Exported {len(product_data)} products to Excel.")

if __name__ == "__main__":
    main()
