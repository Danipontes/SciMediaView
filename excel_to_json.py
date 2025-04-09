import pandas as pd
import json

def excel_to_json(input_file, output_file):
    # Carregar a planilha Excel
    df = pd.read_excel(input_file)
    
    # Renomear colunas para formato mais amigável
    df.columns = [col.strip() for col in df.columns]
    df = df.rename(columns={
        'divisão': 'divisao',
        'descdivisao': 'descdivisao',
        'grupo': 'grupo',
        'descgrupo': 'descgrupo',
        'News_mentions': 'News_mentions',
        'Blog_mentions': 'Blog_mentions',
        'Policy_mentions': 'Policy_mentions',
        'Patent_mentions': 'Patent_mentions',
        'X_mentions': 'X_mentions',
        'Peer_review_mentions': 'Peer_review_mentions',
        'Weibo_mentions': 'Weibo_mentions',
        'Facebook_mentions': 'Facebook_mentions',
        'Wikipedia_mentions': 'Wikipedia_mentions',
        'GooglePlus_mentions': 'GooglePlus_mentions',
        'LinkedIn_mentions': 'LinkedIn_mentions',
        'Reddit_mentions': 'Reddit_mentions',
        'Pinterest_mentions': 'Pinterest_mentions',
        'F1000_mentions': 'F1000_mentions',
        'Q_A_mentions': 'QA_mentions',
        'Video_mentions': 'Video_mentions',
        'Syllabi_mentions': 'Syllabi_mentions',
        'Number_of_Mendeley_readers': 'Mendeley_readers',
        'Number_of_Dimensions_citations': 'Dimensions_citations'
    })
    
    # Converter para JSON
    data = df.to_dict(orient='records')
    
    # Salvar em arquivo JSON
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"Arquivo JSON salvo em {output_file}")

if __name__ == "__main__":
    # Configurações (altere conforme necessário)
    input_excel = r"C:\Users\Ocean\Downloads\dashborad\EXPTESE-00-Outputresearch-altmetric.xlsx"
    output_json = r"C:\Users\Ocean\Downloads\dashborad\data.json"
    
    excel_to_json(input_excel, output_json)