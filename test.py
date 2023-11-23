
import re
text = "Python是一种广泛使用的高级编程语言，用于Web开发、数据分析、Python人工智能等领域。"
pattern = "Python"

matches = re.finditer(pattern, text)

for match in matches:
    txt_match = match.group()
    print("匹配到的文本：", match.group())
    print("起始位置：", match.start())
    print("结束位置：", match.end())
    print()


def split_list(list1, list2):
    positions = [list2.index(i) for i in list1 if i in list2]
    lists = [list2[positions[i]:positions[i+1]] for i in range(len(positions)-1)]
    lists.append(list2[positions[-1]:])
    return lists,positions

list1 = ['a', 'h', 'j']
list2 = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
print(split_list(list1, list2))