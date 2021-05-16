from .models import Room
from django.urls import path
from .views import JoinRoom, RoomView, CreateRoomView, GetRoom

urlpatterns = [
    path('home', RoomView.as_view()),
    path('create-room', CreateRoomView.as_view()),
    path("get-room", GetRoom.as_view()),
    path("join-room", JoinRoom.as_view())
]
