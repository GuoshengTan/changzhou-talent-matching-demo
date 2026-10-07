"""仅用标准库运行。能力检索演示，不进行录用或求职资格判定。"""
import json,math,re,hashlib,csv
from collections import Counter
from pathlib import Path
BASE=Path(__file__).resolve().parent
def grams(text):
 text=re.sub(r'[^\u4e00-\u9fffA-Za-z0-9]','',text).lower()
 return Counter(text[i:i+2] for i in range(len(text)-1))
def run(data,alpha=.7):
 cs=data['candidates']; bags=[grams(c['summary']) for c in cs];df=Counter(k for b in bags for k in b)
 idf={k:math.log((1+len(cs))/(1+n))+1 for k,n in df.items()}
 def vector(text):
  v={k:(1+math.log(n))*idf[k] for k,n in grams(text).items() if k in idf};z=math.sqrt(sum(x*x for x in v.values()))
  return {k:x/z for k,x in v.items()} if z else {}
 vv=[vector(c['summary']) for c in cs];result={}
 for q in data['queries']:
  qv=vector(q['text']);rows=[]
  for c,cv in zip(cs,vv):
   hit=sorted(set(q['skills'])&set(c['skills']));missing=sorted(set(q['skills'])-set(c['skills']))
   coverage=len(hit)/len(q['skills']);sim=sum(x*cv.get(k,0) for k,x in qv.items())
   rows.append(dict(candidate_id=c['id'],name=c['name'],coverage=coverage,text_similarity=sim,score=alpha*coverage+(1-alpha)*sim,covered=hit,not_evidenced=missing,eligibility='未核验',cooperation_interest='未知'))
  result[q['id']]=sorted(rows,key=lambda r:(-r['score'],r['candidate_id']))
 return result
def main():
 data=json.loads((BASE/'public_data.json').read_text(encoding='utf8'));rankings=run(data)
 sensitivity={str(a):{q:[r['candidate_id'] for r in rows[:3]] for q,rows in run(data,a).items()} for a in [.5,.7,.9]}
 out={'method':'0.7×显式标签覆盖率 + 0.3×字符二元组TF-IDF余弦相似度','weights_status':'人为设置，未经真实标签训练','data_sha256':hashlib.sha256((BASE/'public_data.json').read_bytes()).hexdigest(),'rankings':rankings,'sensitivity_top3':sensitivity,'evaluation':None,'evaluation_reason':'没有独立匹配标注或真实对接结果，不计算准确率、NDCG或录用率。','cf_enabled':False,'bert_enabled':False}
 (BASE/'results.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf8')
 with (BASE/'全部匹配结果.csv').open('w',encoding='utf-8-sig',newline='') as f:
  w=csv.writer(f);w.writerow(['需求方向','排序','人才编号','姓名','演示分数','标签覆盖率','文本相似度','命中标签','尚无证据标签'])
  for q,rows in rankings.items():
   for i,r in enumerate(rows):w.writerow([q,i+1,r['candidate_id'],r['name'],round(r['score'],6),round(r['coverage'],6),round(r['text_similarity'],6),'；'.join(r['covered']),'；'.join(r['not_evidenced'])])
 for q,rows in rankings.items():print(q,[(r['name'],round(r['score'],3)) for r in rows[:3]])
if __name__=='__main__':main()
