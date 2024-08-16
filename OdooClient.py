"""
The OdooClient class is a wrapper for the XML-RPC protocol used to access data from an Odoo server.
It is design to simplify the common queries used to create, read, update, and delete records from Odoo models.
"""
import xmlrpc.client

class OdooClient:
    def __init__(self, url, db, username, password, uid=False):
        self.url = url
        self.db = db
        self.username = username
        self.password = password

        self.common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        if not uid:
            self.uid = self.common.authenticate(db, username, password, {})
        else:
            self.uid = uid
        self.models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

    """
    :return a dictionary of the fields in the specified model in the following format: 
        {field: {'type': 'description of type', 'help': 'description or other help about the field if available', 'string': 'name of field'}}
    :param modelName: name of the model to get fields from
    :param fields: list of model fields to get info
    :param attrs: list of attributes to get about the requested fields.
    """
    def getFields(self, modelName, fields=None, attrs=None):
        if attrs is None:
            attrs = []

        if fields is None:
            fields = []
        else:
            fields = [fields]
        return self.models.execute_kw(
            self.db, self.uid, self.password, modelName, 'fields_get',
            fields, {'attributes': attrs})

    """
    :return the ids of the models matching the specified field conditions
    :param modelName: name of the model to search
    :param fieldConditions: list of field conditions in the following format: [['fieldName', '=', value],['fieldName', '=', value]]
    :param limit: number of desired records (int)
    """
    def search(self, modelName, fieldConditions, limit=None):
       if limit is None:
           return self.models.execute_kw(self.db, self.uid, self.password,
            modelName, 'search',
            [fieldConditions])
       return self.models.execute_kw(self.db, self.uid, self.password,
            modelName, 'search',
            [fieldConditions],
            {'limit': limit})

    """
    :return the specified fields from the specified ids
    :param modelName: name of the model to search
    :param ids: list of ids to read.
    :param fields: list of desired fields.    
    """
    def read(self, modelName, ids, fields=None):
        if not isinstance(ids, list):
            ids = [ids]
        if fields is None:
            return self.models.execute_kw(self.db, self.uid, self.password,
            modelName, 'read', [ids])

        else:
            return self.models.execute_kw(self.db, self.uid, self.password,
                modelName, 'read',
                [ids], {'fields': fields})

    """ 
    :return the specified fields from the ids that meet the specified field conditions
    :param modelName: name of the model to search
    :param fieldConditions: list of field conditions in the following format [['fieldName', '=', value],['fieldName', '=', value]]
    :param fields: list of desired fields.
    :param limit: number of desired records (int)
    """
    def searchRead(self, modelName, fieldConditions, fields=None, limit=None):
        ids = self.search(modelName,fieldConditions,limit)
        return self.read(modelName,ids,fields)

    """
    creates new record in the specified model.
    :param modelName: name of the model in which to create a new record.
    :param data: dictionary of fields and their desired value.
    """
    def create(self, modelName, data, options=None):
        if options is None:
            options = {}
        return self.models.execute_kw(self.db, self.uid, self.password, modelName, 'create', [data], options)

    """
    updates specified record in the specified model
    :param modelName: name of the desired model
    :param ID: id of the record to update or list of ids
    :param data: dictionary of fields and their desired value
    """
    def update(self, modelName, ID, data, options=None):
        if options is None:
            options = {}
        if not isinstance(ID, list):
            ID = [ID]
        return self.models.execute_kw(self.db, self.uid, self.password, modelName, 'write', [ID, data], options)