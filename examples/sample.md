#[Sales]
@freeze(1)
@filter

```table
Region,Revenue,Cost,Margin
North America,$1200,$840,30%
Europe,$1450,$1030,29%
Asia Pacific,$1380,$980,29%
Latin America,$1600,$1090,32%
```

| **Total** | $$=SUM(B2:B5)$$ | $$=SUM(C2:C5)$$ | $$=AVERAGE(D2:D5)$$ |

#[Tasks]
@style(A=6 center, B=32, C=12, D=14 lightgreen)

| ID | Task | Owner | Status |
|----|------|-------|--------|
| 1 | {bg:lightyellow}API design | Maria | {dropdown:todo,doing,done}done |
| 2 | **Database schema** | {fg:blue}Kenji | {dropdown:todo,doing,done}doing |
| 3 | {comment:starts 2024-02-01}Integration tests | Aisha | {dropdown:todo,doing,done}todo |
| 4 | Set up `ci.yml` | {align:right}Lucas | {dropdown:todo,doing,done}todo |

See the [design guide](https://example.com/guide) before picking up a task.

#[Spec]
@document

# 1. Functional requirements
## 1.1 Sign-in
Users authenticate with an email address and a password.
- Enter email address
  - must be a valid address
- Enter password
  - masked input
- Verify credentials
## 1.2 Dashboard
The first screen after sign-in.
- Show KPI summary
- Show recent notifications
# 2. Non-functional requirements
## 2.1 Performance
1. Response time: under 3 seconds
2. Concurrent users: 1,000
## 2.2 Security
- TLS required
- Session timeout: 30 minutes
## 2.3 Localization
- All UI strings translatable
- Dates and currencies follow the user's locale

#[Milestones]

## Launch plan

| Milestone | City | Date | Budget |
|-----------|------|------|--------|
| Kickoff | New York | 2024-01-15 | $5,000 |
| Beta review | Tokyo | 2024-03-20 | ¥150,000 |
| Launch event | Berlin | 2024-06-01 | $12,000 |

---
*Budgets are in local currency.*
