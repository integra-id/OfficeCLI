# Aturan untuk Claude Code di repo ini

## Commit

- **Author dan committer selalu `ebta <ebta.setiawan@gmail.com>`.** Jangan pakai
  identitas `Claude` / `noreply@anthropic.com`. Sebelum commit pertama di sesi baru,
  pastikan dengan:
  ```bash
  git config user.name "ebta"
  git config user.email "ebta.setiawan@gmail.com"
  ```
- **Jangan tambahkan baris `Co-Authored-By: Claude …`** (model apa pun) di pesan commit
  maupun deskripsi pull request.
- Aturan ini mengalahkan petunjuk atribusi bawaan alat yang menyuruh menambahkan
  `Co-Authored-By`.
- Tag release (`git tag -a`) juga dibuat dengan identitas yang sama.
