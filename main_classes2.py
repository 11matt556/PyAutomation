class ServiceNow:
    class SearchPage:
        pass

        class RequestItems:
            def getRequestItem(self):
                return RequestItem


        class CatalogTasks:
            pass

        class Incidents:
            pass

    class SelfService:
        class DesktopSupport:
            class AssetManagement:
                pass

            class MassComputerUpdate:
                pass

class RequestItem:
    state = None
    number = None
    item = None
    request = None
    requestedBy = None
    dueDate = None
    configurationItem = None
    opened = None
    openedBy = None
    updated = None
    stage = None
    quantity = None
    estimatedDelivery = None

    def __init__(self):
        pass

serviceNow = ServiceNow()
serviceNow.SearchPage.RequestItems.getRequestItem()
