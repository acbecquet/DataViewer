from win32com.client import gencache
gencache.is_readonly = False
gen_py_dir = gencache.GetGeneratePath()
for root, dirs, files in os.walk(gen_py_dir):
    for file in files:
        os.unlink(os.path.join(root, file))
    for dir in dirs:
        os.rmdir(os.path.join(root, dir))
print("Cleared `gen_py` cache.")


gencache.EnsureDispatch("Excel.Application")
print("Generated `gen_py` files.")