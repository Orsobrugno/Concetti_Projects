import spacy
import pandas as pd
import re


def create_nlp_pipeline(code_list):
    nlp = spacy.blank ( "en" )
    ruler = nlp.add_pipe ( "entity_ruler", config={"overwrite_ents": True} )
    patterns = [{"label": "CONCODE", "pattern": code} for code in code_list]
    ruler.add_patterns ( patterns )
    return nlp


def extract_codes_and_quantities(nlp, text):
    code_quantity = []

    # Split text into lines
    lines = text.strip ().split ( '\n' )

    # Iterate over each line
    for line in lines:
        # Use regex to find quantities
        quantity_match = re.search ( r"(\d+,\d+|\d+\.?\d*)", line )
        if quantity_match:
            quantity = float ( quantity_match.group ().replace ( ',', '.' ) )

            # Use NLP to find codes
            doc = nlp ( line )
            for ent in doc.ents:
                if ent.label_ == "CONCODE":
                    code_quantity.append ( (ent.text, quantity) )

    return code_quantity


def format_output(code_quantity):
    formatted_output = []
    for code, qty in code_quantity:
        formatted_output.append ( f"{code}\t{qty:.2f}" )
    return formatted_output


def save_to_excel(code_quantity, output_file):
    df = pd.DataFrame ( code_quantity, columns=["Code", "Quantity"] )
    df.to_excel ( output_file, index=False )


# Load the list of codes from a file
with open ( "codes.txt" ) as f:
    codes = [line.strip () for line in f]

nlp = create_nlp_pipeline ( codes )

# Example input text
text = """
Material\tTexto breve\tQtd.solicitada
5000196618\tCORREIA W001TR08L50575 CONCETTI\t4,00
5000209674\tEIXO 11014065A-CONCETTI\t1,00
5000209675\tMANOPLA 0001TR07H2024-CONCETTI\t1,00
5000209674\tEIXO 11014065A-CONCETTI\t2,00
5000209675\tMANOPLA 0001TR07H2024-CONCETTI\t2,00
5000209404\tMOLA 0002MO0110478 CONCETTI\t2,00
5000209579\tVIGA HA844/001 CONCETTI\t1,00
5000209580\tVIGA HA844003 CONCETTI\t1,00
5000209404\tMOLA 0002MO0110478 CONCETTI\t1,00
5000197573\tROLO W001TR14R18601 CONCETTI\t3,00
5000197574\tROLO W001TR14R18603 CONCETTI\t3,00
5000197576\tRODA DENTADA X3058000 CONCETTI\t6,00
5000197350\tAVANCO ESQUERDO MOLA 11011449M CONCETTI\t27,00
5000197351\tAVANCO DIREITA MOLA 11011450M CONCETTI\t27,00
5000203233\tCREMALHEIRA X3068000 CONCETTI\t2,00
"""

# Extract codes and quantities
code_quantity = extract_codes_and_quantities ( nlp, text )

# Format the output
formatted_output = format_output ( code_quantity )

# Print the output
print ( "Output for the input text:" )
for line in formatted_output:
    print ( line )

# Save the results to an Excel file
save_to_excel ( code_quantity, "output.xlsx" )
