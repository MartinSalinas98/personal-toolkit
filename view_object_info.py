import pandas as pd

df = pd.DataFrame([[10, 20, 30], [100, 200, 300]], columns=['foo', 'bar', 'baz'])

def get_methods(object, spacing=20):
    """ Display class methods """
    methodList = []
    for method_name in dir(object):
        try:
            if callable(getattr(object, method_name)):
                methodList.append(str(method_name))
        except Exception:
            methodList.append(str(method_name))
    processFunc = (lambda s: ' '.join(s.split())) or (lambda s: s)
    for method in methodList:
        try:
            print(str(method.ljust(spacing)) + ' ' +
                processFunc(str(getattr(object, method).__doc__)[0:90]))
        except Exception:
            print(method.ljust(spacing) + ' ' + ' getattr() failed')

def get_attributes(object):
    """ Display class attributes. """
    try:
        attributes = [attr for attr in dir(object) if not callable(getattr(object, attr, None)) and not attr.startswith("__")]
        print("\n".join(attributes))
    except AttributeError:
        print(type(object), "is not a class instance.")

def display_info(object, infoType='all'):
    """ Displays selected info. """
    if infoType == 'all':
        display_info(object, infoType='methods')
        display_info(object, infoType='attributes')
    elif infoType == 'methods':
        print("METHODS:")
        get_methods(object)
    elif infoType == 'attributes':
        print("ATTRIBUTES:")
        get_attributes(object)