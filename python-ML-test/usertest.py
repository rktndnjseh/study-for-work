# 사용자 정보를 저장할 리스트
users = []

# 사용자 추가 (Create)
def create_user(name, age):
    users.append({'name': name, 'age': age})
    print(f"{name}님이 추가되었습니다.")

# 사용자 목록 보기 (Read)
def read_users():
    if not users:
        print("등록된 사용자가 없습니다.")
    for user in users:
        print(f"이름: {user['name']}, 나이: {user['age']}")

# 사용자 정보 수정 (Update)
def update_user(name, new_age):
    for user in users:
        if user['name'] == name:
            user['age'] = new_age
            print(f"{name}님의 나이가 {new_age}세로 수정되었습니다.")
            return
    print(f"{name}님을 찾을 수 없습니다.")

# 사용자 삭제 (Delete)
def delete_user(name):
    for user in users:
        if user['name'] == name:
            users.remove(user)
            print(f"{name}님이 삭제되었습니다.")
            return
    print(f"{name}님을 찾을 수 없습니다.")

# 테스트
create_user("철수", 20)
create_user("영희", 25)
read_users()
update_user("철수", 21)
delete_user("영희")
read_users()
