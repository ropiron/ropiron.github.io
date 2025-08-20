#!/bin/sh
# ==== 設定（必要に応じて書き換え） ==========================================
USER_NAME="sanei"
SSH_PORT="10022"
PUBLIC_KEY='ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAINSS47CLwmmXSZ0bwEoLMASBC7IR+RZ7yBoPDaB2X2Oo'

# ==== 0) 便利オプション =======================================================
set -eu

# ログに時刻を残す
log(){ printf '%s %s\n' "$(date +'%F %T')" "$*"; }

# ==== 1) タイムゾーン/基本パッケージ ==========================================
log "[1] timezone & base packages"
timedatectl set-timezone Asia/Tokyo 2>/dev/null || true

apt-get update -y
DEBIAN_FRONTEND=noninteractive apt-get upgrade -y
apt-get install -y ca-certificates curl gnupg ufw fail2ban git \
                   python3-venv python3-pip

# ==== 2) 一般ユーザー作成 & SSH鍵配置 =========================================
log "[2] create user & authorized_keys"
if ! id "$USER_NAME" >/dev/null 2>&1; then
  adduser --disabled-password --gecos "" "$USER_NAME"
  usermod -aG sudo "$USER_NAME"
fi

home_dir="$(getent passwd "$USER_NAME" | cut -d: -f6)"
install -d -m 700 "$home_dir/.ssh"
printf '%s\n' "$PUBLIC_KEY" > "$home_dir/.ssh/authorized_keys"
chmod 600 "$home_dir/.ssh/authorized_keys"
chown -R "$USER_NAME:$USER_NAME" "$home_dir/.ssh"

# ==== 3) SSH ハードニング（Port/Password/Root） ===============================
log "[3] harden sshd_config (port=$SSH_PORT)"
SSHD="/etc/ssh/sshd_config"
cp -p "$SSHD" "${SSHD}.$(date +%m%d-%H%M%S).bak" || true

# 既存ディレクティブを書き換え（無ければ追記）
ensure_kv () {
  key="$1"; val="$2"
  if grep -Eiq "^[#[:space:]]*$key[[:space:]]+" "$SSHD"; then
    sed -i "s~^[#[:space:]]*${key}[[:space:]].*~${key} ${val}~I" "$SSHD"
  else
    printf '%s %s\n' "$key" "$val" >> "$SSHD"
  fi
}
ensure_kv "Port" "$SSH_PORT"
ensure_kv "PasswordAuthentication" "no"
ensure_kv "PermitRootLogin" "no"
ensure_kv "UseDNS" "no"
ensure_kv "PubkeyAuthentication" "yes"

# 再起動（ssh.socket がある環境も考慮）
systemctl restart sshd || systemctl restart ssh || true
systemctl list-unit-files | grep -q "^ssh.socket" && systemctl restart ssh.socket || true

# ==== 4) UFW（22 を閉じて 10022/80/443 許可） ===============================
log "[4] configure UFW"
ufw --force reset
ufw default deny incoming
ufw default allow outgoing
ufw allow "${SSH_PORT}/tcp"
ufw allow 80/tcp
ufw allow 443/tcp
ufw --force enable

# ==== 5) fail2ban（最低限のsshd保護） ========================================
log "[5] enable fail2ban (sshd)"
cat >/etc/fail2ban/jail.d/sshd.local <<EOF
[sshd]
enabled = true
port    = ${SSH_PORT}
filter  = sshd
logpath = /var/log/auth.log
maxretry = 5
bantime  = 1h
findtime = 10m
EOF
systemctl enable --now fail2ban

# ==== 6) Docker公式リポジトリ → Docker & Compose v2 ==========================
log "[6] install Docker & Compose v2"
install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg \
 | gpg --dearmor -o /etc/apt/keyrings/docker.gpg
chmod a+r /etc/apt/keyrings/docker.gpg

codename="$(. /etc/os-release && echo "$VERSION_CODENAME")"
echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] \
https://download.docker.com/linux/ubuntu ${codename} stable" \
> /etc/apt/sources.list.d/docker.list

apt-get update -y
apt-get install -y docker-ce docker-ce-cli containerd.io \
                   docker-buildx-plugin docker-compose-plugin
systemctl enable --now docker

# 一般ユーザーを docker グループへ（要再ログインで反映）
usermod -aG docker "$USER_NAME"

# ==== 7) 作業用ディレクトリ（任意） ==========================================
log "[7] prepare /srv/app"
mkdir -p /srv/app
chown -R "$USER_NAME:$USER_NAME" /srv/app

log "[DONE] SSH:${SSH_PORT}, UFW enabled, fail2ban ON, Docker ready."
