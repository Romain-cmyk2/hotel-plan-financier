"""
Synchronisation GitHub pour persistance des plans.

Streamlit Community Cloud utilise un filesystem ephemere : les fichiers ecrits
en local par l'app sont perdus a chaque redemarrage du conteneur. Pour persister
les plans, on pousse chaque sauvegarde sur le depot GitHub via l'API Contents.

Configuration via st.secrets (sur Streamlit Cloud) :
    GITHUB_TOKEN  = "github_pat_xxx"  (fine-grained, scope Contents R/W)
    GITHUB_REPO   = "owner/repo"
    GITHUB_BRANCH = "main"            (optionnel, defaut: main)

En local (sans secrets), is_enabled() renvoie False et les fonctions sont des no-op.
"""

import base64
import time
import requests
import streamlit as st


API_BASE = "https://api.github.com"
TIMEOUT = 15
MIN_INTERVAL_SEC = 5.0  # Throttle: minimum 5s entre 2 pushs sur le meme fichier


def _config():
    """Lit la config depuis st.secrets. Renvoie None si non configure."""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets["GITHUB_REPO"]
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        return token, repo, branch
    except (KeyError, FileNotFoundError, AttributeError):
        return None


def is_enabled():
    return _config() is not None


def _headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _get_file(repo, branch, path, token):
    """Recupere (content_str, sha) ou (None, None) si inexistant/erreur."""
    url = f"{API_BASE}/repos/{repo}/contents/{path}"
    try:
        r = requests.get(url, headers=_headers(token), params={"ref": branch}, timeout=TIMEOUT)
    except requests.RequestException:
        return None, None
    if r.status_code != 200:
        return None, None
    j = r.json()
    try:
        content = base64.b64decode(j["content"]).decode("utf-8")
        return content, j.get("sha")
    except (KeyError, ValueError):
        return None, None


def push_file(path_in_repo, content_str, commit_message):
    """
    Cree ou met a jour un fichier sur GitHub.
    - Throttle: au plus 1 push toutes les MIN_INTERVAL_SEC pour le meme fichier.
    - Verifie cote serveur si le contenu est deja identique.
    - Retry une fois sur 409 avec un SHA frais.
    Renvoie (ok: bool, message: str).
    """
    cfg = _config()
    if cfg is None:
        return False, "GitHub sync non configure (mode local)"
    token, repo, branch = cfg

    # Throttle par fichier via st.session_state (par session Streamlit)
    throttle_key = f"_gh_last_push_ts::{path_in_repo}"
    now = time.time()
    last_ts = st.session_state.get(throttle_key, 0)
    if now - last_ts < MIN_INTERVAL_SEC:
        return True, f"Throttle: dernier push il y a {now - last_ts:.1f}s"

    # GET le contenu et le SHA actuels juste avant le PUT.
    # Si identique cote serveur, on ne pousse pas.
    current_content, sha = _get_file(repo, branch, path_in_repo, token)
    if current_content is not None and current_content == content_str:
        st.session_state[throttle_key] = now
        return True, "Deja a jour sur GitHub"

    payload = {
        "message": commit_message,
        "content": base64.b64encode(content_str.encode("utf-8")).decode("ascii"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    url = f"{API_BASE}/repos/{repo}/contents/{path_in_repo}"
    try:
        r = requests.put(url, headers=_headers(token), json=payload, timeout=TIMEOUT)
    except requests.RequestException as e:
        return False, f"Reseau: {e}"

    # En cas de conflit (409), retry une fois avec un SHA frais.
    if r.status_code == 409:
        _, fresh_sha = _get_file(repo, branch, path_in_repo, token)
        if fresh_sha:
            payload["sha"] = fresh_sha
            try:
                r = requests.put(url, headers=_headers(token), json=payload, timeout=TIMEOUT)
            except requests.RequestException as e:
                return False, f"Reseau (retry): {e}"

    if r.status_code in (200, 201):
        st.session_state[throttle_key] = time.time()
        return True, "Synchronise sur GitHub"
    return False, f"GitHub HTTP {r.status_code}: {r.text[:150]}"


def pull_file(path_in_repo):
    """
    Recupere le contenu live d'un fichier sur GitHub.
    Renvoie (content_str, sha) ou (None, None).
    """
    cfg = _config()
    if cfg is None:
        return None, None
    token, repo, branch = cfg

    url = f"{API_BASE}/repos/{repo}/contents/{path_in_repo}"
    try:
        r = requests.get(url, headers=_headers(token), params={"ref": branch}, timeout=TIMEOUT)
    except requests.RequestException:
        return None, None
    if r.status_code != 200:
        return None, None
    j = r.json()
    try:
        content = base64.b64decode(j["content"]).decode("utf-8")
    except (KeyError, ValueError):
        return None, None
    return content, j.get("sha")


def list_files(directory):
    """
    Liste les fichiers JSON dans un repertoire du repo.
    Renvoie une liste de noms (sans extension), ou None si non configure / erreur.
    """
    cfg = _config()
    if cfg is None:
        return None
    token, repo, branch = cfg

    url = f"{API_BASE}/repos/{repo}/contents/{directory}"
    try:
        r = requests.get(url, headers=_headers(token), params={"ref": branch}, timeout=TIMEOUT)
    except requests.RequestException:
        return None
    if r.status_code != 200:
        return None
    items = r.json()
    if not isinstance(items, list):
        return None
    return [
        item["name"][:-5]
        for item in items
        if item.get("name", "").endswith(".json")
    ]


def delete_file(path_in_repo, commit_message):
    """Supprime un fichier sur GitHub. Renvoie (ok, message)."""
    cfg = _config()
    if cfg is None:
        return False, "GitHub sync non configure"
    token, repo, branch = cfg

    _, sha = _get_file(repo, branch, path_in_repo, token)
    if sha is None:
        return True, "Deja absent sur GitHub"

    url = f"{API_BASE}/repos/{repo}/contents/{path_in_repo}"
    try:
        r = requests.delete(
            url,
            headers=_headers(token),
            json={"message": commit_message, "sha": sha, "branch": branch},
            timeout=TIMEOUT,
        )
    except requests.RequestException as e:
        return False, f"Reseau: {e}"
    if r.status_code == 200:
        return True, "Supprime sur GitHub"
    return False, f"GitHub HTTP {r.status_code}: {r.text[:150]}"


def rename_file(old_path, new_path, content_str, commit_message):
    """Renomme : push sous le nouveau chemin puis supprime l'ancien."""
    ok, msg = push_file(new_path, content_str, commit_message)
    if not ok:
        return False, msg
    ok2, msg2 = delete_file(old_path, commit_message)
    if not ok2:
        return False, f"Cree mais ancien non supprime : {msg2}"
    return True, "Renomme sur GitHub"


def sync_directory_from_github(local_dir, remote_dir):
    """
    Synchronise un repertoire local depuis GitHub : telecharge tous les fichiers
    JSON manquants ou potentiellement obsoletes. A appeler au demarrage.
    Renvoie le nombre de fichiers telecharges.
    """
    cfg = _config()
    if cfg is None:
        return 0
    names = list_files(remote_dir)
    if not names:
        return 0
    count = 0
    for name in names:
        content, _ = pull_file(f"{remote_dir}/{name}.json")
        if content is None:
            continue
        local_path = local_dir / f"{name}.json"
        try:
            existing = local_path.read_text(encoding="utf-8") if local_path.exists() else None
        except OSError:
            existing = None
        if existing != content:
            local_path.write_text(content, encoding="utf-8")
            count += 1
    return count
