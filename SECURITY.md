# セキュリティポリシー

## 脆弱性の報告

脆弱性を見つけた場合は、**公開 Issue には書かず**、GitHub の [Private vulnerability reporting](https://github.com/Miyanber/JRWTimeTableGenerator/security/advisories/new) から非公開で報告してください。

報告には、可能な範囲で次の内容を含めてください。

- 脆弱性の概要と影響
- 再現手順（使用したファイルや環境）
- 考えられる修正方法（任意）

本プロジェクトは個人が趣味で開発・保守しているため、対応時期はお約束できません。

## 対象範囲

- このリポジトリのソースコード（`JRWTimeTableGeneratorWeb/`、マニフェスト）
- GitHub Actions のワークフロー

外部サービス（JRおでかけネット、Google Apps Script、Microsoft の CDN）そのものは対象外です。
