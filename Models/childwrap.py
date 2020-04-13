class child_wrap_obj:
     def __init__(self, objName, isExists, fieldWrapperList, parentObjWrapperList, QueryCondition):
         self.objName = objName
         self.isExists = isExists
         self.fieldWrapperList = fieldWrapperList
         self.parentObjWrapperList =parentObjWrapperList
         self.QueryCondition= QueryCondition