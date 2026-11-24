// OfficeSDK mark uses the circular artwork from icon.svg cropped to its painted bounds.
export const officeIcon = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjcgNyA1MCA1MCI+PGNpcmNsZSBjeD0iMzIiIGN5PSIzMiIgcj0iMjUiIGZpbGw9IiMwMDAiLz48cGF0aCB0cmFuc2Zvcm09Im1hdHJpeCgzLjU3MTQgMCAwIDMuNTcxNCAtMy43MTQgLTMuNzE0KSIgZD0iTTkuNDkzOTcgMTIuMDI0MUM5Ljc1MzU5IDEzLjI0IDEwLjY2MjMgMTMuNjM2OCAxMS42ODY3IDEyLjc4MzFDMTIuMTkyOCAxMi4zNjE0IDEyLjUwOTYgMTEuOTYwNiAxMi42OSAxMS42NDM1QzEzLjI4OTIgMTAuNTkwNCAxMy4xOTM0IDkuODU3NTcgMTIuMzYxNCA5LjA3MjI4QzExLjMyNjEgOC4wOTQ5NSA5LjgxMzMxIDcuMjQ2MzEgNy4xNTAwMiA2LjM4OTM2QzYuOTE0MDYgNi4zMTY0MyA2LjgxNzg2IDYuNTA5MzkgNi45MjAwNyA2LjY3MTk3QzguMDMzNzggNy41MTIyIDguODUzNzYgOS4wMjU1OSA5LjQ5Mzk3IDEyLjAyNDFaIiBmaWxsPSIjZmZmIi8+PC9zdmc+';
export const officeWebsite = 'https://officesdk.com/';

export function openOfficeWebsite(doc) {
  const link = Object.assign(doc.createElement('a'), { href:officeWebsite, target:'_blank', rel:'noopener noreferrer' });
  doc.body.appendChild(link);
  link.click();
  link.remove();
}
