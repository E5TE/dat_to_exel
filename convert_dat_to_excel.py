import pandas as pd
import os
import struct

def try_different_encodings(dat_file):
    encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']
    
    for encoding in encodings:
        try:
            with open(dat_file, 'rb') as f:
                # Try to read the binary data
                data = f.read()
                
                # Try to decode with current encoding
                text = data.decode(encoding)
                
                # Try different delimiters
                for delimiter in ['\t', ',', '|', ';']:
                    try:
                        # Convert the text to a DataFrame
                        from io import StringIO
                        df = pd.read_csv(StringIO(text), delimiter=delimiter)
                        if len(df.columns) > 1:  # Check if we got meaningful columns
                            return df
                    except:
                        continue
                
                # If no delimiter worked, try fixed width
                try:
                    df = pd.read_fwf(StringIO(text))
                    if len(df.columns) > 1:
                        return df
                except:
                    continue
                    
        except:
            continue
    
    # If no text-based approach worked, try binary approach
    try:
        with open(dat_file, 'rb') as f:
            # Read binary data in chunks and try to interpret as numbers
            data = []
            while True:
                try:
                    # Try reading 4 bytes as a float
                    chunk = f.read(4)
                    if not chunk:
                        break
                    if len(chunk) == 4:
                        value = struct.unpack('f', chunk)[0]
                        data.append(value)
                except:
                    break
            
            if data:
                return pd.DataFrame(data, columns=['Value'])
    except:
        pass
    
    raise Exception("Could not parse the DAT file with any known method")

def convert_dat_to_excel(dat_file):
    try:
        print(f"Attempting to convert {dat_file}...")
        
        # Try to read the data with different methods
        df = try_different_encodings(dat_file)
        
        # Create Excel filename
        excel_file = os.path.splitext(dat_file)[0] + '.xlsx'
        
        # Convert to Excel
        df.to_excel(excel_file, index=False)
        print(f"Successfully converted {dat_file} to {excel_file}")
        
    except Exception as e:
        print(f"Error converting file: {str(e)}")
        print("Please ensure the .dat file contains valid data in a structured format.")

if __name__ == "__main__":
    # Convert all .dat files in the current directory
    for file in os.listdir('.'):
        if file.endswith('.dat'):
            convert_dat_to_excel(file)
