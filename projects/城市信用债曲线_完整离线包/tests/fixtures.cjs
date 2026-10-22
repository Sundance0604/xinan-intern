'use strict';
function fixture(date = '2026-10-22') {
  const fields = ['code','name','issuer','term','yield','rating','implied','yy','method','province','city','type','sheet','row','amount','couponText','start','maturity','frequency','windLGFV','yyLGFV','fullName','face'];
  const rows = [2,1,1.2,3,3.2].map((term, i) => ['TEST00' + (i + 1) + '.IB', i ? '验证样本' + i : '更新验证债', '合成验证发行人', term, 1.5 + i / 10, 'AA+', 'AA', '6', '公募', '江苏省', '南京市', '中期票据', '合成数据', i + 2, 1, '2.00%', '2025-01-01', '2029-01-01', '每年付息1次', '是', '是', '合成验证债券', 100]);
  return { meta: { fields, date, filename: '合成验证数据.xlsx', sheets: [{ name: '合成数据', count: 5 }] }, rows };
}
const htmlOf = data => '<html><script>window.CREDIT_DATA=' + JSON.stringify(data).replace(/<\//g, '<\\/') + ';</script><script>globalThis.SHOULD_NOT_RUN=true;document.body.textContent="DO NOT EXECUTE IMPORTED HTML";</script></html>';
module.exports = { fixture, htmlOf };
