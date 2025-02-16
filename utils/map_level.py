def map_level(score) :
    if 0 <= score <= 20:
        return 'FOUNDATION'
    elif 21 <= score <= 40:
        return 'EMERGING'
    elif 41 <= score <= 60:
        return 'PROFICIENT'
    elif 61 <= score <= 80:
        return 'ADVANCED'
    else:
        return 'EXPERT'