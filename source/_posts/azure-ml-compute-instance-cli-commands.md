---
title: Azure ML Compute Instance を CLI で操作する（開始・停止・再起動）
date: 2026-03-26
tags:
  - Azure
  - AzureML
  - CLI
---

Azure Machine Learning の Compute Instance を Azure CLI で開始・停止・再起動する方法をまとめます。

## 対象リソース

| 項目 | 値 |
|------|-----|
| サブスクリプション ID | `1ed842ac-f4bf-46ce-8fc4-591b9edb213b` |
| リソースグループ | `rg-openai-sdv-sid1` |
| ワークスペース名 | `ml-knowledge-reuse20260206` |
| Compute Instance 名 | `CompInsKnowledge20260206` |

## 前提条件

Azure CLI および `ml` 拡張機能がインストール済みであること。

```bash
# Azure CLI のインストール（未インストールの場合）
# https://docs.microsoft.com/ja-jp/cli/azure/install-azure-cli

# ml 拡張機能のインストール
az extension add -n ml

# Azure へのログイン
az login

# サブスクリプションの設定
az account set --subscription 1ed842ac-f4bf-46ce-8fc4-591b9edb213b
```

## Compute Instance の操作コマンド

### 開始（Start）

```bash
az ml compute start \
  --name CompInsKnowledge20260206 \
  --workspace-name ml-knowledge-reuse20260206 \
  --resource-group rg-openai-sdv-sid1 \
  --subscription 1ed842ac-f4bf-46ce-8fc4-591b9edb213b
```

### 停止（Stop）

```bash
az ml compute stop \
  --name CompInsKnowledge20260206 \
  --workspace-name ml-knowledge-reuse20260206 \
  --resource-group rg-openai-sdv-sid1 \
  --subscription 1ed842ac-f4bf-46ce-8fc4-591b9edb213b
```

### 再起動（Restart）

```bash
az ml compute restart \
  --name CompInsKnowledge20260206 \
  --workspace-name ml-knowledge-reuse20260206 \
  --resource-group rg-openai-sdv-sid1 \
  --subscription 1ed842ac-f4bf-46ce-8fc4-591b9edb213b
```

## 状態の確認

```bash
az ml compute show \
  --name CompInsKnowledge20260206 \
  --workspace-name ml-knowledge-reuse20260206 \
  --resource-group rg-openai-sdv-sid1 \
  --subscription 1ed842ac-f4bf-46ce-8fc4-591b9edb213b \
  --query "state" \
  --output tsv
```

`state` の値は以下のいずれかになります。

| 状態 | 説明 |
|------|------|
| `Running` | 実行中 |
| `Stopped` | 停止中 |
| `Starting` | 起動処理中 |
| `Stopping` | 停止処理中 |

## 参考リンク

- [az ml compute — Azure CLI リファレンス](https://learn.microsoft.com/ja-jp/cli/azure/ml/compute)
- [Azure Machine Learning コンピューティング インスタンスの管理](https://learn.microsoft.com/ja-jp/azure/machine-learning/how-to-create-manage-compute-instance)
