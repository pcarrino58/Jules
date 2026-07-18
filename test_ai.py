import re
from user_ai_bridge import get_generalized_signature, clean_contractor_noise

print(get_generalized_signature("unit2a"))
print(get_generalized_signature("pump 2a"))
print(get_generalized_signature("pump 1 system"))
print(get_generalized_signature("chiller (water)"))

print(clean_contractor_noise("chiller (water)"))
