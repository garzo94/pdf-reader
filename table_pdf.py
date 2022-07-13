import camelot as cam





class Table_to_pdf:

    def __init__(self, pdf_path):
        self.pdf = pdf_path


    def extract_table(self):
        extract = cam.read_pdf(self.pdf, page=1, flavor='stream')
        df = extract[0].df
        df.drop(labels=3, axis=1, inplace=True)
        columns = df.iloc[1]
        df.columns = columns
        df.drop(labels=[0, 1], inplace=True)
        df.set_index(keys='VARIABLES', inplace=True)
        df.drop('', inplace=True)
        df.columns.name = None
        df.replace(to_replace=['kgf', 'g', 'mm', 'mL', '/-', r'\+', ' '], value='', regex=True, inplace=True)
        df.ESPECIFICACIÓN.iloc[:].astype(float).astype(int)
        df.TOLERANCIA.iloc[0:5].astype(str).astype(int)
        df.TOLERANCIA.iloc[7:12].astype(str).astype(int)
        return df




