from .gstr3b_master import generate_gstr3b_master
from .gstr1_master import generate_gstr1_master
from .gstr2a_master import generate_gstr2a_master
from .ewb_in_master import generate_ewb_in_master
from .ewb_out_master import generate_ewb_out_master

async def generate_master_excel_for_return_type(return_type, input_dir, output_dir):
    if return_type == "GSTR-3B":
        return await generate_gstr3b_master(input_dir, output_dir)
    elif return_type == "GSTR-1":
        return await generate_gstr1_master(input_dir, output_dir)
    elif return_type == "GSTR-2A":
        return await generate_gstr2a_master(input_dir, output_dir)
    elif return_type == "EWB-IN":
        return await generate_ewb_in_master(input_dir, output_dir)
    elif return_type == "EWB-OUT":
        return await generate_ewb_out_master(input_dir, output_dir)
    else:
        raise ValueError(f"Unsupported return type: {return_type}")
