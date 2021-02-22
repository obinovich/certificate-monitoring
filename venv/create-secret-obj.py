import pickle

#sample to enter the credentials "field:value"
#example_dict = {'sl_api_key':'xxxxxxxxxxxxxxxxx'}
dict = {'uname':'xxx', 'pwd':'xxx'}

# specify filename for secret eg "secrets.obj"
pickle_out = open("xxxoff-secrets.obj","wb")
pickle.dump(dict, pickle_out)
pickle_out.close()
