if __name__ == '__main__':
    final_list = []

    for n in range(int(input())):
        temp_list = []
        name = input()
        temp_list.append(name)
        score = float(input())
        temp_list.append(score)
        final_list.append(temp_list)
    final_list.sort(key=lambda x: x[1])
    print(final_list)



