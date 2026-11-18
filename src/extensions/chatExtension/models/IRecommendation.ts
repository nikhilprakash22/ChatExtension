export interface IRecommendation {
  id:                string;
  title:             string;
  category:          string;
  description:       string;
  distanceFromHotel: number;
  walkingDistance:   boolean;
  eveningFriendly:   boolean;
  mapLink:           string;
  website:           string;
  isActive:          boolean;
}