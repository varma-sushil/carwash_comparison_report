import openpyxl
import logging
logger =logging.getLogger(__file__)

def check_zero_values(filename, sheet_name):
    # Load the existing workbook using openpyxl
    ret = False
    try:
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook[sheet_name]
        # Iterate over all rows and columns
        for row in worksheet.iter_rows():
            for cell in row:
                value = cell.value
                # print(value==0)
                if value==0:
                    print(cell.coordinate)
                    ret=True
                    break
                # if value is isinstance(value,(int,float)):
                #     if value==0:
                        # print("Error in report")
                
    except KeyError:
        logger.error(f"Sheet not found: {sheet_name}")
        ret=True
    except FileNotFoundError:
        logger.error(f"File not found: {filename}")
        ret=True
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        ret=True
        
    return ret

if __name__ == "__main__":
    # Set logging level to INFO
    logging.basicConfig(level=logging.INFO)
    
    print(check_zero_values(filename="2024.xlsx", sheet_name="2024-08-0"))
