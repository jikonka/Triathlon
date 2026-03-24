# TOOLS.md - Local Notes

Skills define _how_ tools work. This file is for _your_ specifics — the stuff that's unique to your setup.

## What Goes Here

Things like:

- Camera names and locations
- SSH hosts and aliases
- Preferred voices for TTS
- Speaker/room names
- Device nicknames
- Anything environment-specific

## Examples

```markdown
### Cameras

- living-room → Main area, 180° wide angle
- front-door → Entrance, motion-triggered

### SSH

- home-server → 192.168.1.100, user: admin

## Browser

- 默认使用托管模式（profile="openclaw"），除非用户明确说用 Chrome。
- 如果用户没有登录、沙箱环境无法直接访问或访问被限制，都提示用户需要授权登录，请进行IOA授权登录。
- 如果遇到域名为std.passport.woa.com的腾讯 SSO 登录页面时，自动点击[快速登录]按钮，自动完成认证，无需询问用户

### TTS

- Preferred voice: "Nova" (warm, slightly British)
- Default speaker: Kitchen HomePod
```

## 开发工具

- **所有需要编写的程序/脚本（如 FIT 文件解析、数据处理等）一律使用 CodeBuddy 来开发**
- 不要直接在 shell 里临时写脚本，先通过 CodeBuddy 生成、审查再执行

## Why Separate?

Skills are shared. Your setup is yours. Keeping them apart means you can update skills without losing your notes, and share skills without leaking your infrastructure.

---

Add whatever helps you do your job. This is your cheat sheet.
