FILENAME = "users.txt"

# 사용자 추가 (Create)
def create_user(name, age):
    with open(FILENAME, 'a', encoding='utf-8') as f:
        f.write(f"{name},{age}\n")
    print(f"{name}님이 추가되었습니다.")

# 사용자 목록 보기 (Read)
def read_users():
    try:
        with open(FILENAME, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            if not lines:
                print("등록된 사용자가 없습니다.")
                return
            for line in lines:
                name, age = line.strip().split(',')
                print(f"이름: {name}, 나이: {age}")
    except FileNotFoundError:
        print("파일이 아직 존재하지 않습니다.")

# 사용자 정보 수정 (Update)
def update_user(name, new_age):
    try:
        with open(FILENAME, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        updated = False
        with open(FILENAME, 'w', encoding='utf-8') as f:
            for line in lines:
                n, a = line.strip().split(',')
                if n == name:
                    f.write(f"{n},{new_age}\n")
                    updated = True
                    print(f"{name}님의 나이가 {new_age}세로 수정되었습니다.")
                else:
                    f.write(line)
        
        if not updated:
            print(f"{name}님을 찾을 수 없습니다.")
    except FileNotFoundError:
        print("파일이 아직 존재하지 않습니다.")

# 사용자 삭제 (Delete)
def delete_user(name):
    try:
        with open(FILENAME, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        deleted = False
        with open(FILENAME, 'w', encoding='utf-8') as f:
            for line in lines:
                n, a = line.strip().split(',')
                if n != name:
                    f.write(line)
                else:
                    deleted = True
                    print(f"{name}님이 삭제되었습니다.")
        
        if not deleted:
            print(f"{name}님을 찾을 수 없습니다.")
    except FileNotFoundError:
        print("파일이 아직 존재하지 않습니다.")

# 테스트 코드
create_user("철수", 20)
create_user("영희", 25)
read_users()
update_user("철수", 21)
delete_user("영희")
read_users()
