"""
Generate a DIM wishlist from the Monument of Triumph tab in the weapon rolls spreadsheet.
"""

import json
import openpyxl
from collections import defaultdict

ITEMS_FILE = r"C:\Users\james\AppData\Local\Temp\destiny_items.json"
PLUGSETS_FILE = r"C:\Users\james\AppData\Local\Temp\destiny_plugsets.json"
XLSX_FILE = r"C:\Users\james\Desktop\destiny-2-wishlists\Destiny 2 Weapon Rolls.xlsx"
OUTPUT_FILE = r"C:\Users\james\Desktop\destiny-2-wishlists\monument_of_triumph_wishlist.txt"
SHEET_NAME = "Monument of Triumph"

# Weapon name typo corrections from the spreadsheet
WEAPON_CORRECTIONS = {
    "Alone as a God": "Alone as a god",
    "Accured Redemption": "Accrued Redemption",
    "Judgement": "Judgment",
    "1000-Yard Stare": "1000 Yard Stare",
    "Matador-64": "Matador 64",
    "Vouchesafe": "Vouchsafe",
    "Twighlight Oath": "Twilight Oath",
    "Bonechiler": "Bonechiller",
    "Thermal Errosion": "Thermal Erosion",
    "Orign Story": "Origin Story",
    "Plug One.1": "PLUG ONE.1",
    "OGMA PR6": "Ogma PR6",
    "Snipehunt MK.47": "Snipehunt Mk. 47",
    "Accelerando-47": "Accelerando-42",
    "Terciopelo-4BL": "Terciopelo-4bl",
    "Fair Judgement": "Fair Judgment",
    "Prolonged Engamenet": "Prolonged Engagement",
    "NOX Perrenial V": "Nox Perennial V",
    "Tempation's Hook": "Temptation's Hook",
    "Heliocentric QSC": "Heliocentric QSc",
    "Inbound Surveilance": "Inbound Surveillance",
    "Compase Rose": "Compass Rose",
}

# Perk name typo corrections from the spreadsheet
PERK_CORRECTIONS = {
    "Flutted Barrel": "Fluted Barrel",
    "Bewildering Blast": "Bewildering Burst",
    "Crystaline Corpsebloom": "Crystalline Corpsebloom",
    "Crystalline Corpsbloom": "Crystalline Corpsebloom",
    "Impromptu Ammunition": "Impromptu Ammo",
    "Corkscrew": "Corkscrew Rifling",
    "Auxillary Reserves": "Auxiliary Reserves",
    "Cluster Bombs": "Cluster Bomb",
    "Snapshot": "Snapshot Sights",
    "High Velocity Rounds": "High-Velocity Rounds",
    "Tacttical Mag": "Tactical Mag",
    "Moving Taregt": "Moving Target",
    "Resevoir Burst": "Reservoir Burst",
    "Fourth Times the Charm": "Fourth Time's the Charm",
    "Subsistance": "Subsistence",
    "Distruption Break": "Disruption Break",
    "Chain reaction": "Chain Reaction",
    "Low-Impedance Wendings": "Low-Impedance Windings",
    "Accurized": "Accurized Rounds",
    "High-Explosive Ordinance": "High-Explosive Ordnance",
    "Handd": None,  # Masterwork typo, not a perk - skip
}

# Perk name alternatives (if the primary name isn't found on the weapon)
PERK_ALTERNATIVES = {
    "Impromptu Ammunition": ["Impromptu Ammo"],
    "Impromptu Ammo": ["Impromptu Ammunition"],
    "Bewildering Burst": ["Bewildering Blast"],
    "Crystalline Corpsebloom": ["Crystaline Corpsebloom", "Crystalline Corpsbloom"],
    "Auxiliary Reserves": ["Auxillary Reserves"],
    "Fluted Barrel": ["Flutted Barrel"],
    "Snapshot Sights": ["Snapshot"],
    "Cluster Bomb": ["Cluster Bombs"],
    "High-Velocity Rounds": ["High Velocity Rounds"],
    "Tactical Mag": ["Tacttical Mag"],
    "Moving Target": ["Moving Taregt"],
    "Reservoir Burst": ["Resevoir Burst"],
    "Fourth Time's the Charm": ["Fourth Times the Charm"],
    "Subsistence": ["Subsistance"],
    "Disruption Break": ["Distruption Break"],
    "Chain Reaction": ["Chain reaction"],
    "Low-Impedance Windings": ["Low-Impedance Wendings"],
}

def load_manifests():
    print("Loading item manifest...")
    with open(ITEMS_FILE, "r", encoding="utf-8") as f:
        items = json.load(f)

    print("Loading plugset manifest...")
    with open(PLUGSETS_FILE, "r", encoding="utf-8") as f:
        plugsets = json.load(f)

    return items, plugsets

def build_lookup_tables(items, plugsets):
    print("Building lookup tables...")

    # hash -> item name
    hash_to_name = {}
    # hash -> tierType (2=Common/regular, 3=Uncommon/enhanced, 5=Legendary/old)
    hash_to_tier = {}
    # perk name -> list of (hash, tierType) sorted to prefer regular first
    perk_name_to_hashes = defaultdict(list)
    # weapon name -> list of weapon data dicts
    weapon_name_to_list = defaultdict(list)

    for h_str, item in items.items():
        h = int(h_str)
        dp = item.get("displayProperties", {})
        name = dp.get("name", "").strip()
        if not name:
            continue

        hash_to_name[h] = name
        item_type = item.get("itemType", 0)
        tier = item.get("inventory", {}).get("tierType", 0)
        hash_to_tier[h] = tier

        # Collect plug items (perks, barrels, mags, etc.)
        if item.get("plug"):
            perk_name_to_hashes[name].append((h, tier))

        # Collect weapons (itemType 3)
        if item_type == 3:
            sockets = item.get("sockets", {}).get("socketEntries", [])
            # Require at least 5 sockets to be a real weapon
            if len(sockets) >= 5:
                weapon_name_to_list[name].append({
                    "hash": h,
                    "sockets": sockets,
                    "socket_count": len(sockets),
                    "item": item,
                })

    # Sort each perk's hashes: prefer regular (tierType=2) over old (5) over enhanced (3)
    for name in perk_name_to_hashes:
        perk_name_to_hashes[name].sort(key=lambda x: (0 if x[1] == 2 else 1 if x[1] == 5 else 2))

    # Build set of enhanced perk hashes (tierType=3) to avoid in wishlist
    enhanced_hashes = {h for hlist in perk_name_to_hashes.values() for h, t in hlist if t == 3}

    # plugSet hash -> plug hashes sorted so non-enhanced appear before enhanced
    plugset_to_plug_hashes = {}
    for ps_h_str, ps_data in plugsets.items():
        ps_h = int(ps_h_str)
        all_plugs = [p["plugItemHash"] for p in ps_data.get("reusablePlugItems", [])]
        # Non-enhanced first so find_perk_hash_on_weapon returns them preferentially
        all_plugs.sort(key=lambda ph: (1 if ph in enhanced_hashes else 0))
        plugset_to_plug_hashes[ps_h] = all_plugs

    print(f"  Weapons: {len(weapon_name_to_list)} unique names")
    print(f"  Perk names: {len(perk_name_to_hashes)}")
    print(f"  PlugSets: {len(plugset_to_plug_hashes)}")
    print(f"  Enhanced perk hashes: {len(enhanced_hashes)}")

    return hash_to_name, hash_to_tier, perk_name_to_hashes, weapon_name_to_list, plugset_to_plug_hashes, enhanced_hashes

def pick_best_weapon(candidates):
    """Pick the weapon candidate with the most sockets (most complete definition)."""
    if not candidates:
        return None
    return max(candidates, key=lambda w: w["socket_count"])

def find_weapon_with_best_perk_match(candidates, perk_names, hash_to_name, plugset_to_plug_hashes,
                                      perk_name_to_hashes, enhanced_hashes):
    """Try all weapon candidates and return the one that resolves the most perks."""
    best_candidate = None
    best_resolved = -1

    for candidate in candidates:
        resolved = 0
        for perk_name in perk_names:
            if not perk_name:
                continue
            h = find_perk_hash_on_weapon(perk_name, candidate, hash_to_name,
                                          plugset_to_plug_hashes, perk_name_to_hashes,
                                          enhanced_hashes, range(1, 8))
            if h is not None:
                resolved += 1
        if resolved > best_resolved:
            best_resolved = resolved
            best_candidate = candidate

    # Fall back to highest socket count if no perks resolved for any
    if best_candidate is None:
        best_candidate = pick_best_weapon(candidates)
    return best_candidate

def get_all_plugs_in_socket(sock, plugset_to_plug_hashes):
    """Get all plug item hashes available in a socket."""
    plugs = set()

    # From randomized plug set
    rsp = sock.get("randomizedPlugSetHash")
    if rsp and rsp in plugset_to_plug_hashes:
        plugs.update(plugset_to_plug_hashes[rsp])

    # From reusable plug set
    rp = sock.get("reusablePlugSetHash")
    if rp and rp in plugset_to_plug_hashes:
        plugs.update(plugset_to_plug_hashes[rp])

    # From single initial item hash
    sp = sock.get("singleInitialItemHash")
    if sp:
        plugs.add(sp)

    return plugs

def find_perk_hash_on_weapon(perk_name, weapon_data, hash_to_name, plugset_to_plug_hashes,
                              perk_name_to_hashes, enhanced_hashes, socket_indices=range(1, 7)):
    """Find the hash of a named perk in a weapon's sockets.
    Always returns the non-enhanced (regular) hash when both exist.
    Returns the perk hash, or None if not found.
    """
    sockets = weapon_data["sockets"]

    # Collect candidate hashes for this perk name (non-enhanced preferred via sorted order)
    candidate_hashes = {h for h, _ in perk_name_to_hashes.get(perk_name, [])}

    # Also add alternatives
    for alt in PERK_ALTERNATIVES.get(perk_name, []):
        candidate_hashes.update(h for h, _ in perk_name_to_hashes.get(alt, []))

    for sock_idx in socket_indices:
        if sock_idx >= len(sockets):
            continue
        sock = sockets[sock_idx]
        all_plugs = get_all_plugs_in_socket(sock, plugset_to_plug_hashes)

        # Among matching plugs, prefer non-enhanced (not in enhanced_hashes)
        matches = [ph for ph in all_plugs if ph in candidate_hashes or hash_to_name.get(ph, "") == perk_name]
        if not matches:
            continue
        # Return non-enhanced version if available, else fall back to any match
        regular = [ph for ph in matches if ph not in enhanced_hashes]
        return regular[0] if regular else matches[0]

    return None

def find_perk_hash_globally(perk_name, perk_name_to_hashes, enhanced_hashes):
    """Fallback: return first non-enhanced known hash for a perk name."""
    hashes = perk_name_to_hashes.get(perk_name, [])
    if not hashes:
        return None
    # Prefer non-enhanced (already sorted: regular first, then enhanced)
    for h, tier in hashes:
        if h not in enhanced_hashes:
            return h
    return hashes[0][0]

def correct_perk_name(name):
    """Apply typo corrections to perk names. Handles combined perks (e.g. 'Perk A + Perk B')."""
    if name is None:
        return None
    # Handle combined perk entries like "Attrition Orbs + Vorpal Weapon" - take the first one
    if " + " in name:
        name = name.split(" + ")[0].strip()
    corrected = PERK_CORRECTIONS.get(name, name)
    return corrected

def correct_weapon_name(name):
    """Apply typo corrections to weapon names."""
    if name is None:
        return None
    return WEAPON_CORRECTIONS.get(name, name)

def load_spreadsheet_data():
    print("Loading spreadsheet...")
    wb = openpyxl.load_workbook(XLSX_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    entries = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        weapon_name = row[0]
        if weapon_name is None:
            continue
        source = row[1] or ""
        use = row[2] or ""
        col1_perk = correct_perk_name(row[4])   # Barrel
        col2_perk = correct_perk_name(row[6])   # Magazine
        col3_perk = correct_perk_name(row[8])   # Perk 1
        col4_perk = correct_perk_name(row[10])  # Perk 2
        masterwork = row[12] or ""
        mod = row[13] or ""
        origin_trait = row[14] or ""

        entries.append({
            "weapon_name": correct_weapon_name(weapon_name.strip()),
            "source": source.strip(),
            "use": use.strip(),
            "barrel": col1_perk,
            "magazine": col2_perk,
            "perk1": col3_perk,
            "perk2": col4_perk,
            "masterwork": masterwork,
            "mod": mod,
            "origin_trait": origin_trait,
        })

    print(f"Loaded {len(entries)} weapon entries")
    return entries

def generate_wishlist(entries, hash_to_name, perk_name_to_hashes, weapon_name_to_list,
                      plugset_to_plug_hashes, enhanced_hashes):
    lines = []

    lines.append("title:Monument of Triumph God Rolls")
    lines.append("description:Curated weapon rolls for Monument of Triumph weapons. "
                 "Based on the Destiny 2 Weapon Rolls spreadsheet by kimo23.")
    lines.append("")

    not_found_weapons = []
    not_found_perks = []
    processed = 0

    for entry in entries:
        weapon_name = entry["weapon_name"]
        source = entry["source"]
        use = entry["use"]

        # Find weapon in manifest - try all candidates and pick best perk match
        candidates = weapon_name_to_list.get(weapon_name, [])
        perk_names_for_matching = [entry["barrel"], entry["magazine"], entry["perk1"], entry["perk2"]]
        weapon_data = find_weapon_with_best_perk_match(
            candidates, perk_names_for_matching, hash_to_name, plugset_to_plug_hashes,
            perk_name_to_hashes, enhanced_hashes
        )

        if not weapon_data:
            not_found_weapons.append(f"{weapon_name} ({source})")
            # Still write a commented-out entry
            lines.append(f"// WEAPON NOT FOUND: {weapon_name} | {source} | {use}")
            lines.append(f"// Barrel: {entry['barrel']} | Mag: {entry['magazine']} | "
                         f"Perk1: {entry['perk1']} | Perk2: {entry['perk2']}")
            lines.append("")
            continue

        weapon_hash = weapon_data["hash"]

        # Resolve perk hashes - search within weapon's sockets first
        perk_data = {
            "barrel": entry["barrel"],
            "magazine": entry["magazine"],
            "perk1": entry["perk1"],
            "perk2": entry["perk2"],
        }

        # Socket index ranges for each perk type
        socket_search = {
            "barrel": list(range(1, 4)),
            "magazine": list(range(1, 4)),
            "perk1": list(range(1, 7)),
            "perk2": list(range(1, 7)),
        }

        resolved_perks = {}
        missing_perks = []

        for perk_key, perk_name in perk_data.items():
            if not perk_name:
                continue
            # Search within weapon sockets
            perk_hash = find_perk_hash_on_weapon(
                perk_name, weapon_data, hash_to_name, plugset_to_plug_hashes,
                perk_name_to_hashes, enhanced_hashes, socket_search[perk_key]
            )
            if perk_hash is None:
                # Fallback: global search (non-enhanced preferred)
                perk_hash = find_perk_hash_globally(perk_name, perk_name_to_hashes, enhanced_hashes)

            if perk_hash is not None:
                resolved_perks[perk_key] = perk_hash
            else:
                missing_perks.append(f"{perk_key}={perk_name}")

        if missing_perks:
            not_found_perks.append(f"{weapon_name}: {', '.join(missing_perks)}")

        # Build perk hash list (barrel, magazine, perk1, perk2)
        perk_hashes = []
        for perk_key in ["barrel", "magazine", "perk1", "perk2"]:
            if perk_key in resolved_perks:
                perk_hashes.append(resolved_perks[perk_key])

        # Build notes
        notes_parts = [
            f"{weapon_name}",
            f"Source: {source}",
            f"Use: {use}",
        ]
        if entry["masterwork"]:
            notes_parts.append(f"MW: {entry['masterwork']}")
        if entry["mod"]:
            notes_parts.append(f"Mod: {entry['mod']}")
        if entry["origin_trait"]:
            notes_parts.append(f"Origin: {entry['origin_trait']}")
        perk_names_str = " / ".join(filter(None, [
            entry["barrel"], entry["magazine"], entry["perk1"], entry["perk2"]
        ]))
        notes_parts.append(f"Perks: {perk_names_str}")

        tag = "pve" if "PVE" in use.upper() else "pvp"
        notes_line = " | ".join(notes_parts)

        # Write comment header
        lines.append(f"// {weapon_name} - {source} - {use}")
        lines.append(f"// Perks: {entry['barrel']} / {entry['magazine']} / "
                     f"{entry['perk1']} / {entry['perk2']}")
        if entry["masterwork"]:
            lines.append(f"// MW: {entry['masterwork']} | Mod: {entry['mod']}")

        if perk_hashes:
            lines.append(f"//notes:{notes_line}|tags:{tag}")
            perks_str = ",".join(str(h) for h in perk_hashes)
            lines.append(f"dimwishlist:item={weapon_hash}&perks={perks_str}")
            # Also add a "any roll" entry with just the key perks (perk1, perk2)
            key_perks = [resolved_perks.get("perk1"), resolved_perks.get("perk2")]
            key_perks = [h for h in key_perks if h is not None]
            if len(key_perks) == 2 and key_perks != perk_hashes[-2:]:
                lines.append(f"// Also matches without specific barrel/mag:")
                lines.append(f"dimwishlist:item={weapon_hash}&perks={','.join(str(h) for h in key_perks)}")
        else:
            lines.append(f"// Could not resolve perk hashes for this weapon")

        lines.append("")
        processed += 1

    return lines, not_found_weapons, not_found_perks, processed

def main():
    items, plugsets = load_manifests()
    hash_to_name, hash_to_tier, perk_name_to_hashes, weapon_name_to_list, plugset_to_plug_hashes, enhanced_hashes = build_lookup_tables(items, plugsets)
    entries = load_spreadsheet_data()

    print(f"\nGenerating wishlist for {len(entries)} entries...")
    lines, not_found_weapons, not_found_perks, processed = generate_wishlist(
        entries, hash_to_name, perk_name_to_hashes, weapon_name_to_list,
        plugset_to_plug_hashes, enhanced_hashes
    )

    # Write output
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\nWishlist written to: {OUTPUT_FILE}")
    print(f"Processed: {processed} weapons")
    if not_found_weapons:
        print(f"\nWeapons NOT found in manifest ({len(not_found_weapons)}):")
        for w in not_found_weapons:
            print(f"  - {w}")
    if not_found_perks:
        print(f"\nPerks NOT resolved ({len(not_found_perks)}):")
        for p in not_found_perks[:20]:
            print(f"  - {p}")
        if len(not_found_perks) > 20:
            print(f"  ... and {len(not_found_perks) - 20} more")

if __name__ == "__main__":
    main()
