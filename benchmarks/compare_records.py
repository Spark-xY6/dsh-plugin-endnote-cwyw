"""Score automatic duplicate decisions on two pinned LDD bibliographic datasets."""

import argparse
import csv
from hashlib import sha256
from importlib.metadata import version
from io import StringIO
from itertools import combinations
import json
from math import comb
from pathlib import Path
import sys

import pandas as pd
from bib_dedupe import bib_dedupe
from bib_dedupe.cluster import get_connected_components
from tqdm import tqdm

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "skills/cite-from-pdfs/scripts"))
from reference_identity import Identity, compare


FILES = ("records_pre_merged.csv", "merged_record_ids.csv")
HASHES = {
    "stroke": (
        "a1ea76cdfe5c0fa98896fec8fa74d116deac3b20e06a67ba7a288b787c7793ed",
        "f932a243ef78d5909d2a8c23a2db7c440de7b8bf732c80d351366aaec10e0987",
    ),
    "haematology": (
        "c8ec567898eb90ef53bbd9db3924ecbebd4eb6efbe2bca4fa28bda2f893cb8b7",
        "b874b2de2ee7d62a2637817c8b26a77de42fb1fc0bbae4e8fed3de475775a16d",
    ),
}


def pairs(groups, valid_ids):
    result = set()
    for group in groups:
        if not set(group) <= valid_ids or len(group) != len(set(group)):
            raise ValueError("Invalid ID in a predicted or labelled group")
        result.update(combinations(sorted(group), 2))
    return result


def read_gold(contents, valid_ids):
    seen = set()
    groups = []
    with StringIO(contents, newline="") as stream:
        for row in csv.DictReader(stream):
            group = row["merged_ids"].split(";")
            if seen.intersection(group):
                raise ValueError("Overlapping gold clusters")
            seen.update(group)
            groups.append(group)
    return pairs(groups, valid_ids)


def plugin_predictions(records):
    entries = []
    for row in records:
        first_author = row["author"].split(" and ", 1)[0].partition(",")[0]
        identity = Identity.from_record(dict(row, first_author_family=first_author))
        entries.append((row["ID"], identity))
    accepted, review = set(), set()
    for i, (left_id, left) in enumerate(tqdm(entries, desc="Reference records", unit="record")):
        for right_id, right in entries[i + 1:]:
            decision = compare(left, right)
            if decision in {"ambiguous", "conflict"}:
                review.add(tuple(sorted((left_id, right_id))))
            elif decision in {"doi", "pmid", "title_author_year", "title_variant"}:
                accepted.add(tuple(sorted((left_id, right_id))))
            elif decision != "different":
                raise ValueError(f"Unexpected identity decision {decision!r}")
    return accepted, review


def measure(accepted, review, gold, total):
    tp, fp, fn = len(accepted & gold), len(accepted - gold), len(gold - accepted)
    if review & accepted:
        raise ValueError("Review and automatic decisions overlap")
    return {
        "automatic_pairs": len(accepted), "tp": tp, "fp": fp, "fn": fn,
        "tn": total - tp - fp - fn,
        "precision": tp / len(accepted) if accepted else None,
        "recall": tp / len(gold) if gold else None,
        "review_pairs": len(review), "gold_pairs_in_review": len(review & gold),
    }


def score(root, dataset):
    tables = []
    for name, expected in zip(FILES, HASHES[dataset]):
        path = root / dataset / name
        contents = path.read_bytes()
        if sha256(contents).hexdigest() != expected:
            raise ValueError(f"Input checksum mismatch: {path.name} in {dataset}")
        tables.append(contents.decode("utf-8-sig"))
    records_text, labels_text = tables
    with StringIO(records_text, newline="") as stream:
        records = list(csv.DictReader(stream))
    ids = {row["ID"] for row in records}
    if "" in ids or len(ids) != len(records):
        raise ValueError("Empty or duplicate record ID")
    gold = read_gold(labels_text, ids)
    total = comb(len(ids), 2)
    accepted, review = plugin_predictions(records)
    frame = pd.read_csv(StringIO(records_text), dtype=str, keep_default_na=False)
    frame["container_title"] = frame["journal"]
    prepared = bib_dedupe.prep(frame, cpu=1, verbosity_level=0)
    if set(prepared["ID"]) != ids or len(prepared) != len(ids):
        raise ValueError("BibDedupe preparation changed the record population")
    blocked = bib_dedupe.block(prepared, cpu=1, verbosity_level=0)
    matched = bib_dedupe.match(blocked, cpu=1, verbosity_level=0)
    clustered = pairs(get_connected_components(matched), ids)

    def labelled(label):
        selected = matched.loc[matched["duplicate_label"] == label, ["ID_1", "ID_2"]]
        return pairs(selected.itertuples(index=False, name=None), ids)

    direct = labelled("duplicate")
    maybe = labelled("maybe")
    return {
        "records": len(ids), "possible_pairs": total, "gold_pairs": len(gold),
        "source_sha256": dict(zip(FILES, HASHES[dataset])),
        "plugin_pairwise": measure(accepted, review, gold, total),
        "bibdedupe_clusters": measure(clustered, maybe - clustered, gold, total),
        "bibdedupe_direct": measure(direct, maybe - direct, gold, total),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", type=Path, required=True, help="Root with stroke and haematology subfolders")
    parser.add_argument("--output", type=Path, required=True, help="New result JSON path")
    args = parser.parse_args()
    if version("bib-dedupe") != "0.11.0":
        parser.error("This comparison requires bib-dedupe==0.11.0")
    results = {name: score(args.data, name) for name in HASHES}
    report = {"bib_dedupe": version("bib-dedupe"), "pandas": pd.__version__, "datasets": results}
    with args.output.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(report, indent=2, allow_nan=False) + "\n")


if __name__ == "__main__":
    main()
