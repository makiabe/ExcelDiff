"""One-time, checksummed transfer of approved WebP assets. Repository-local only."""
import base64
import hashlib
import itertools
import json
import os
import re
import urllib.request
import zlib


def digest(s):
    return hashlib.sha256(s.encode('ascii')).hexdigest()[:16]


def restore(spec, text, patches):
    normalized = ''.join(c for c, _ in itertools.groupby(text))
    wanted = set(spec['hashes'])
    found = {}
    lengths = {spec['block'], spec['size'] % spec['block'] or spec['block']}
    for length in lengths:
        for i in range(len(normalized) - length + 1):
            part = normalized[i:i + length]
            h = digest(part)
            if h in wanted:
                found[h] = part
    for index, part in patches.items():
        h = spec['hashes'][int(index)]
        if digest(part) != h:
            raise ValueError('Patch checksum mismatch: ' + index)
        found[h] = part
    missing = [i for i, h in enumerate(spec['hashes']) if h not in found]
    if missing:
        return None, missing
    normalized = ''.join(found[h] for h in spec['hashes'])
    if len(normalized) != spec['size']:
        raise ValueError('Normalized length mismatch')
    packed = zlib.decompress(base64.b64decode(spec['runs'], validate=True))
    numbers = []
    value = shift = 0
    for b in packed:
        value |= (b & 127) << shift
        if b & 128:
            shift += 7
        else:
            numbers.append(value)
            value = shift = 0
    if shift or len(numbers) % 2:
        raise ValueError('Invalid run metadata')
    runs = {}
    index = 0
    for delta, length in zip(numbers[0::2], numbers[1::2]):
        index += delta
        runs[index] = length
    original = ''.join(c * runs.get(i, 1) for i, c in enumerate(normalized))
    raw = base64.b64decode(original, validate=True)
    if hashlib.sha256(raw).hexdigest() != spec['sha256']:
        raise ValueError('Final asset checksum mismatch')
    return raw, []


def main():
    repo = os.environ['GITHUB_REPOSITORY']
    if repo != 'makiabe/ExcelDiff':
        raise SystemExit('Restricted to makiabe/ExcelDiff')
    root = 'https://api.github.com/repos/' + repo
    token = os.environ['GH_TOKEN']

    def api(path, data=None, method=None):
        body = None if data is None else json.dumps(data).encode()
        req = urllib.request.Request(root + path, data=body,
            method=method or ('GET' if data is None else 'POST'),
            headers={'Authorization': 'Bearer ' + token,
                     'Accept': 'application/vnd.github+json',
                     'Content-Type': 'application/json',
                     'X-GitHub-Api-Version': '2022-11-28'})
        with urllib.request.urlopen(req, timeout=60) as response:
            return json.load(response)

    def contents(path, ref):
        obj = api('/contents/' + path + '?ref=' + ref)
        return base64.b64decode(obj['content'])

    ref = os.environ['GITHUB_SHA']
    manifest = json.loads(contents('.github/artwork-manifest.json', ref))
    patches = json.loads(contents('.github/artwork-patches.json', ref))
    recovered = {}
    missing = {}
    for name, spec in manifest.items():
        pieces = []
        for source in spec['inputs']:
            blob = api('/git/blobs/' + source['sha'])
            if source['kind'] == 'binary':
                pieces.append(''.join(blob['content'].split()))
            else:
                pieces.append(base64.b64decode(blob['content']).decode('ascii'))
        text = re.sub(r'\s+', '', ''.join(pieces))
        raw, absent = restore(spec, text, patches.get(name, {}))
        if absent:
            missing[name] = absent
        else:
            recovered[name] = raw
            print('Verified:', name, len(raw), spec['sha256'])
    if missing:
        print('MISSING_BLOCKS=' + json.dumps(missing, separators=(',', ':')))
        raise SystemExit('No unverified image was published. Complete the transfer patches.')
    head = api('/git/ref/heads/main')['object']['sha']
    base = api('/git/commits/' + head)['tree']['sha']
    entries = []
    for name, raw in recovered.items():
        blob = api('/git/blobs', {'content': base64.b64encode(raw).decode(), 'encoding': 'base64'})
        entries.append({'path': 'assets/' + name + '.webp', 'mode': '100644', 'type': 'blob', 'sha': blob['sha']})
    html = contents('index.html', head).decode('utf-8')
    for name in recovered:
        html = html.replace('assets/' + name + '.png', 'assets/' + name + '.webp')
    entries.append({'path': 'index.html', 'mode': '100644', 'type': 'blob', 'content': html})
    tree = api('/git/trees', {'base_tree': base, 'tree': entries})
    commit = api('/git/commits', {'message': 'Publish verified, web-optimized cat artwork', 'tree': tree['sha'], 'parents': [head]})
    api('/git/refs/heads/main', {'sha': commit['sha'], 'force': False}, 'PATCH')
    print('PUBLISHED_COMMIT=' + commit['sha'])


if __name__ == '__main__':
    main()
