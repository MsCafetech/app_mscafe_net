# Prisma Access Configuration API（/sse/config/v1）取得可能なオブジェクト一覧

出典：GitHub PaloAltoNetworks/pan.dev の openapi-specs/access/prisma-access-config/*.yaml（2026-10-27 確認）。
GitHub のフォルダ一覧はこの環境から見られないため、ファイル名を推定して存在確認したもの。これ以外にもある可能性あり。
仕様ファイルのコピー：scm-excel/api-spec/

凡例：◎=本ツールで取得（既定○） ○=本ツールに行あり（既定は○なし） －=未対応

## オブジェクト
| 対象 | エンドポイント | 本ツール |
|---|---|---|
| Address / Address Group | addresses / address-groups | ◎ |
| Service / Service Group | services / service-groups | ◎ |
| Application / Group / Filter | applications / application-groups / application-filters | ◎ / ◎ / ○ |
| HIP Object / HIP Profile | hip-objects / hip-profiles | ◎ |
| カスタムURLカテゴリ | url-categories | ◎ |
| URLフィルタリングカテゴリ（定義済み一覧） | url-filtering-categories | － |
| Tag | tags | ○ |
| EDL | external-dynamic-lists | ◎ |
| Schedule | schedules | ○ |
| Region | regions | － |
| Dynamic User Group | dynamic-user-groups | － |
| Auto Tag Action | auto-tag-actions | － |
| 脅威シグネチャ（カスタム） | anti-spyware-signatures / vulnerability-protection-signatures | － |

## セキュリティプロファイル
| 対象 | エンドポイント | 本ツール |
|---|---|---|
| Anti-Spyware / Vulnerability / URL Access / File Blocking / WildFire-AV / DNS Security / Decryption | *-profiles | ◎ |
| Profile Group | profile-groups | ◎ |
| HTTP Header Profile | http-header-profiles | － |

## ルール（position=pre/post）
| 対象 | エンドポイント | 本ツール |
|---|---|---|
| セキュリティルール | security-rules | ◎ |
| 復号ルール / 復号除外 | decryption-rules / decryption-exclusions | ◎ / ◎ |
| 認証ルール | authentication-rules | － |
| App Override ルール | app-override-rules | － |
| QoS ルール / QoS プロファイル | qos-policy-rules / qos-profiles | － |
| トラフィックステアリング | traffic-steering-rules | － |

## 認証・証明書
| 対象 | エンドポイント | 本ツール |
|---|---|---|
| Certificate Profile | certificate-profiles | ○ |
| 証明書 / 信頼CA | certificates / trusted-certificate-authorities | － |
| TLS Service Profile | tls-service-profiles | － |
| 認証プロファイル / 認証ポータル | authentication-profiles / authentication-portals | － |
| LDAP / RADIUS / SAML / TACACS+ / Kerberos サーバプロファイル | *-server-profiles | － |
| ローカルユーザー / グループ | local-users / local-user-groups | － |

## インフラ（RN・SC・GlobalProtect）
| 対象 | エンドポイント | 本ツール |
|---|---|---|
| Remote Network / Service Connection / SCグループ | remote-networks / service-connections / service-connection-groups | － |
| IKE Gateway / IKE Crypto / IPSec Tunnel / IPSec Crypto | ike-gateways / ike-crypto-profiles / ipsec-tunnels / ipsec-crypto-profiles | － |
| BGP ルーティング / 帯域割当 | bgp-routing / bandwidth-allocations | － |
| 内部DNS / 共有インフラ設定 | internal-dns-servers / shared-infrastructure-settings | － |
| GlobalProtect（Mobile Agent）設定 | mobile-agent/agent-profiles 等 | － |
| ロケーション / ライセンス種別 | locations / license-types | － |

## 設定管理
| 対象 | エンドポイント |
|---|---|
| 設定バージョン・Push・ジョブ | config-versions / candidate:push / jobs |
