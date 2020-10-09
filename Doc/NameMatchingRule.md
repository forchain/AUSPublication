

# Matching Rule

## Prepare base table

```flow
st=>start: 开始
io1=>inputoutput: 输入论文表
io2=>inputoutput: 输入作者表

op1=>operation: 查找论文的作者A和机构作者B
op11=>check author , duplicate name, depending on researcher ID

op12=>check paper
cond1=>condition: 两者姓是否相同?
cond2=>condition: 两者名首字母缩写是否相同?
io3=>inputoutput: 保存两者到结果表
io5=>inputoutput: 只保存论文到结果表
cond3=>condition: 查找完毕?
io4=>inputoutput: 输出结果表

e=>end: 结束
st->io1->io2->op1->cond1
cond1(yes)->cond2
cond1(no)->io5
cond2(yes)->io3->cond3
cond2(no)->io5

io5->cond3

cond3(yes)->io4
cond3(no)->op1

io4->e


```





## Calculate Matching Score

```flow
st=>start: Start
io1=>inputoutput: Input a record from base table
pIO=>inputoutput: Paper First Name,Paper Middle Name, Paper Middle Initial
aIO=>inputoutput: Author ID, Author First Name, Author First Initial, Author Middle Name, Author Middle Initial

op1=>operation: Score = 0
cond1=>condition: exists author info
op2=>operation: Score=Score+2 ^ 0



cond1(yes)->op2
cond1(no)->e


io2=>inputoutput: Score
e=>end

st->io1->op1->cond1

cond1->e
```





