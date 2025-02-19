import pandas as pd

def import_txt_data(filename):
    """Reads the .txt file and returns a list of [Y, X, Z] points."""
    data = []
    
    with open(filename, 'r') as file:
        for line in file:
            row = line.strip().split()  # Split values (assumes space-separated)
            if len(row) == 3:  # Ensure three columns exist (Y, X, Z)
                try:
                    y, x, z = float(row[0]), float(row[1]), float(row[2])  # Convert to float
                    data.append([y, x, z])
                except ValueError:
                    print(f"Skipping invalid row: {row}")  # Handle bad rows gracefully
    
    return data

def txt_to_matrices(data, chunk_size=200):
    """Splits data into 200-row chunks and converts each into a 2D matrix."""
    all_matrices = []
    all_y_labels = []
    all_x_labels = []
    chunk = []

    # Separate data every 200 rows
    for i, row in enumerate(data, 1):
        chunk.append(row)
        if i % chunk_size == 0:
            matrix, y_labels, x_labels = process_chunk(chunk)
            all_matrices.append(matrix)
            all_y_labels.append(y_labels)
            all_x_labels.append(x_labels)
            chunk = []

    # Process the last chunk if it isn’t empty
    if chunk:
        matrix, y_labels, x_labels = process_chunk(chunk)
        all_matrices.append(matrix)
        all_y_labels.append(y_labels)
        all_x_labels.append(x_labels)

    return all_matrices, all_y_labels, all_x_labels

def process_chunk(data):
    """Converts a chunk of Y, X, Z data into a 2D matrix format."""
    # Extract unique Y and X values for matrix indexing
    unique_Y = sorted(set(row[0] for row in data))
    unique_X = sorted(set(row[1] for row in data))
    
    # Create index mappings
    y_index_map = {val: i for i, val in enumerate(unique_Y)}
    x_index_map = {val: i for i, val in enumerate(unique_X)}

    # Initialize matrix with zeros
    matrix = [[0] * len(unique_X) for _ in range(len(unique_Y))]

    # Fill matrix with Z values
    for y, x, z in data:
        row_idx = y_index_map[y]
        col_idx = x_index_map[x]
        matrix[row_idx][col_idx] = z

    return matrix, unique_Y, unique_X

def save_to_single_excel_sheet(matrices, y_labels_list, x_labels_list, filename="data_iv_1.xlsx"):
    """Saves all matrices into a single Excel sheet WITHOUT extra blank rows."""
    all_data = []
    
    # Add X labels only once at the top
    all_data.append([""] + list(x_labels_list[0]))  # First row: X labels
    
    for i, (matrix, y_labels) in enumerate(zip(matrices, y_labels_list)):
        # Add matrix data with Y labels
        for y_label, row in zip(y_labels, matrix):
            all_data.append([y_label] + list(row))

    # Convert list to DataFrame and save to Excel
    df = pd.DataFrame(all_data)
    
    with pd.ExcelWriter(filename) as writer:
        df.to_excel(writer, sheet_name="All_Matrices", index=False, header=False)
    
    print(f"All matrices saved in {filename} in a single sheet without blank rows.")

# === RUNNING THE CODE ===
filename = r'C:\Users\LEE JUN SEOP\Desktop\Optima schola\data_iv_meas_20250213(iter=11)\data_iv_1.txt'  # Change this to your actual file

# Step 1: Import Data
data_points = import_txt_data(filename)

# Step 2: Convert to Matrices
matrices, y_labels_list, x_labels_list = txt_to_matrices(data_points)

# Step 3: Save to a Single Excel Sheet (NO BLANK ROWS)
save_to_single_excel_sheet(matrices, y_labels_list, x_labels_list)

