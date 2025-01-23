import os
import sys
import ifcopenshell
import ifcopenshell.util.element
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

def generate_xlsx_filename(filename):
    path, ext = os.path.splitext(filename)
    return path + '.xlsx'

def ifc_open(filename):
    return ifcopenshell.open(filename)

def ifc_classes(ifc_file, parent_type='IfcProduct'):
    classes = ifc_file.by_type(parent_type)
    class_names = [class_name.is_a() for class_name in classes]
    class_names = list(set(class_names))
    # class_names.sort()
    return class_names

def ifc_properties_dict(ifc_file, class_list):
    result_dict = {}
    for class_name in class_list:
            objects = ifc_file.by_type(class_name)
            class_arr = []
            for obj in objects:
                class_data = {}
                class_data['Name'] = obj.Name
                class_data['GlobalId'] = obj.GlobalId
                class_data['Description'] = obj.Description
                class_data['ObjectType'] = obj.ObjectType
                class_data['IfcType'] = obj.is_a()
                
                psets = ifcopenshell.util.element.get_psets(obj)
                
                for name, value in psets.items():
                    if isinstance(value, dict):
                        for key, val in value.items():
                            class_data[key] = val
                    else:
                        pass
                class_arr.append(class_data)
            class_df = pd.DataFrame(class_arr)
            if(class_df.empty):
                continue
            result_dict[class_name] = class_df
    return result_dict

def save_properties_to_excel(ifc_filename):
    xlsx_path = generate_xlsx_filename(ifc_filename)
    ifc_file = ifc_open(ifc_filename)
    class_list = ifc_classes(ifc_file)
    ifc_props = ifc_properties_dict(ifc_file, class_list)

    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        for key in ifc_props:
            ifc_props[key].to_excel(writer, sheet_name=key, index=False)

if __name__ == '__main__':
    save_properties_to_excel(sys.argv[1])