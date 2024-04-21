from robocorp.tasks import task
from main import Order
@task
def task():
    order = Order()
    order.get_data()