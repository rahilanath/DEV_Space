import win32com.client
import pythoncom

com_object = win32com.client.Dispatch('Outlook.Application')

for key in dir(com_object):
    method = getattr(com_object,key)

    print(key)
    print(type(method))

    # for key in dir(com_object):
    # method = getattr(com_object,key)
    # if str(type(method)) == "<type 'instance'>":
    #     print(key)
    #     for sub_method in dir(method):
    #         if not sub_method.startswith("_") and not "clsid" in sub_method.lower():
    #             print("\t"+sub_method)
    # else:
    #     print("\t"),method