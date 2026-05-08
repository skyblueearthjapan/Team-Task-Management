# Team Task Management — プロジェクトルール

このファイルは Claude Code が参照する本プロジェクト固有の運用ルールです。

## デプロイメント運用ルール（必須遵守）

作業の節目（設計仕様の確定／実装ステップ完了／レビュー後の修正反映／結合テスト後 等）で以下 2 ステップを必ず実施する。

### 1. GitHub にコミット・プッシュ
- **リポジトリ**: https://github.com/skyblueearthjapan/Team-Task-Management.git
- 対象: ドキュメント・コードを問わずプロジェクト内の変更全般
- 除外: `.env` 等の機密ファイル、`_design_pkg/` 等の一時的な作業フォルダ
- コミットメッセージは HEREDOC 形式で（Claude Code グローバルルール準拠）
- `git push origin main` で本流に反映

### 2. GAS への clasp push
- `src/` 配下に変更があれば実行
- 紐付け先 GAS プロジェクト ID: `1v1P1s5T1L9E7snpsRQT4Wpm2qJTzvt0-kkmgEZ1ec-81lUW0qdBVMlDX`
- 設定ファイル: `.clasp.json`（rootDir = ./src）
- `clasp login` 等の対話操作が必要な場合はユーザーに依頼する

## 実行手順（標準フロー）

```
1. 変更内容のレビュー／QA 完了を確認
2. git status / git diff で差分確認
3. git add <個別ファイル>（-A は使わない）
4. git commit -m "<HEREDOC で記述>"
5. git push origin main
6. （src/ 変更時）clasp push
```

## 重要な注意

- 破壊的操作（force push / main へのリベース等）は事前確認必須
- 機密値（API トークン・パスワード）はコミット禁止
- `.clasp.json` の scriptId は社内専用なので、リポジトリ公開状況に応じて配置を判断する

## 参考メモリ

`~/.claude/projects/.../memory/feedback_deployment_rules.md` に同等のルールを feedback メモリとして保存済み。
