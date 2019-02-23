critics = {'Claudia Puig': {'Just My Luck': 3.0,
  'Snakes on a Plane': 3.5,
  'Superman Returns': 4.0,
  'The Night Listener': 4.5,
  'You, Me, and Dupree': 2.5},
 'Gene Seymour': {'Just My Luck': 1.5,
  'Lady in the Water': 3.0,
  'Snakes on a Plane': 3.5,
  'Superman Returns': 5.0,
  'The Night Listener': 3.0,
  'You, Me, and Dupree': 3.5},
 'Jack Mathews': {'Lady in the Water': 3.0,
  'Snakes on a Plane': 4.0,
  'Superman Returns': 5.0,
  'The Night Listener': 3.0,
  'You, Me, and Dupree': 3.5},
 'Lisa Rose': {'Just My Luck': 3.0,
  'Lady in the Water': 2.5,
  'Snakes on a Plane': 3.5,
  'Superman Returns': 3.5,
  'The Night Listener': 3.0,
  'You, Me, and Dupree': 2.5},
 'Michael Phillips': {'Lady in the Water': 2.5,
  'Snakes on a Plane': 3.0,
  'Superman Returns': 3.5,
  'The Night Listener': 4.0},
 'Mick LaSalle': {'Just My Luck': 2.0,
  'Lady in the Water': 3.0,
  'Snakes on a Plane': 4.0,
  'Superman Returns': 3.0,
  'The Night Listener': 3.0,
  'You, Me, and Dupree': 2.0},
 'Toby': {'Snakes on a Plane': 4.5,
  'Superman Returns': 4.0,
  'You, Me, and Dupree': 1.0},
 'Ali': {'Just My Luck': 2.0,
  'Snakes on a Plane': 2.5,
  'Superman Returns': 2.5,
  'The Night Listener': 5.0}}

from math import sqrt

def sim_distance(prefs, person1, person2):
#This function returns a distance-based similarity score for person1 and person2
#Get the list of shared items between two persons
    si={}
    for item in prefs[person1]:
        if item in prefs[person2]:
            si[item]=1
    if len(si)==0: return 0 #no shared items between two persons
    sum_of_squares = sum([pow(prefs[person1][item] - prefs[person2][item],2)\
                          for item in si])
    return 1/(1+sqrt(sum_of_squares)) #we take a reciprocal to give high score values
#for persons who are similar

def sim_pearson(prefs, p1, p2):
    # This function returns a the Pearson correlation coeffient for p1 and p2
    si = {}
    #First find the items shared by both persons
    for item in prefs[p1]:
        if item in prefs[p2]:
            si[item]=1
    n = len(si)
    if n == 0:
        return 0 #no shared items between two persons
    # Add up all the ratings (i.e. preferences) for shared items per person
    sum1 = sum([prefs[p1][it] for it in si])
    sum2 = sum([prefs[p2][it] for it in si])
    # Addup the squares of all preferences for shared items per person
    sum1Sq = sum([pow(prefs[p1][it],2) for it in si]) 
    sum2Sq = sum([pow(prefs[p2][it],2) for it in si])
    # Addup the cross ratings shared between two persons
    pSum = sum([prefs[p1][it]*prefs[p2][it] for it in si])
    num = pSum-(sum1*sum2/n)
    den = sqrt((sum1Sq-pow(sum1,2)/n)*(sum2Sq-pow(sum2,2)/n))
    # Calculate the Pearson score
    if den == 0:
        return 0
    return num/den


def sim_jaccard(prefs, genre1, genre2):

    #Get the list of items
    genre1_movies = prefs[genre1].keys()
    genre2_movies = prefs[genre2].keys()

    # Make them sets in order to be able to use built-in methods of it such as intersection and union
    p1, p2 = set(genre1_movies), set(genre2_movies)
    p1_intersect_p2 = p1.intersection(p2)
    p1_union_p2 = p1.union(p2)

    #Get the total number of items for intersection and union
    p1_intersect_p2, p1_union_p2 = len(p1_intersect_p2), len(p1_union_p2)

    # if they have no items in common, return 0
    if p1_intersect_p2 == 0: return 0

    return float(p1_intersect_p2) / float(p1_union_p2) # return jaccard distance


def topMatches(prefs, person, n=3, similarity=sim_pearson):
    #This function returns the best matches for person from the prefs dictionary
    scores = [(round(similarity(prefs,person,other), 2),other) \
              for other in prefs if other != person]
    #Sort and reverse the list so that highest scores appears at the top
    scores.sort()
    scores.reverse()
    return scores[0:n]

def getRecommendations(prefs, person, similarity = sim_pearson):
    totals = {}
    simSums = {}
    for other in prefs:#for every other critic
        sim = similarity(prefs, person, other)

        if sim <= 0: continue #go the the next other person if sim<=0
        for item in prefs[other]:#for items reviewed by that critic
            if item not in prefs[person] or prefs[person][item] == 0:#if movie not watched by the person
                totals.setdefault(item, 0) #This method returns the key value available
                #in the dictionary and if given key is not available then it will return provided default value
                totals[item] += prefs[other][item]*sim
                simSums.setdefault(item, 0)
                simSums[item] += sim
    rankings = [(round(total/simSums[item], 2),item) for item, total in totals.items()]
    rankings.sort()
    rankings.reverse()
    return rankings

def transformPrefs(prefs):
    result = {}
    for person in prefs:
	for item in prefs[person]:
        	result.setdefault(item, {})
        	result[item][person] = prefs[person][item]
    return result


def calculateSimilarItems(prefs,n=10):
    # Create a dictionary of items showing which other items they
    # are most similar to.
    result={}
    
    # Invert the preference matrix to be item-centric
    itemPrefs=transformPrefs(prefs)
    c=0
    
    for item in itemPrefs:
        # Status updates for large datasets
        c+=1
        if c%100==0: print "%d / %d" % (c,len(itemPrefs))
        # Find the most similar items to this one
        scores=topMatches(itemPrefs,item,n=n,similarity=sim_distance)
        result[item]=scores

    return result

def getRecommendedItems(prefs,itemMatch,user):
    userRatings=prefs[user]
    scores={}
    totalSim={}

    # Loop over items rated by this user
    for (item,rating) in userRatings.items():
        # Loop over items similar to this one
        for (similarity,item2) in itemMatch[item]:
            # Ignore if this user has already rated this item
            if item2 in userRatings: continue

            # Weighted sum of rating times similarity
            scores.setdefault(item2,0)
            scores[item2]+=similarity*rating

            # Sum of all the similarities
            totalSim.setdefault(item2,0)
            totalSim[item2]+=similarity

    # Divide each total score by total weighting to get an average
    rankings=[(round(score/totalSim[item], 2),item) for item,score in scores.items()]
    # Return the rankings from highest to lowest
    rankings.sort()
    rankings.reverse()
    return rankings
