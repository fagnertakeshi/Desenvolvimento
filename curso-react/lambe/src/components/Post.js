import React , {Component} from 'react';
import { StyleSheet, 
          Text, 
          View,
          Image,
          Dimensions }
           from 'react-native';
import Autor from './Autor'
import Comments from './Comments'


class Post extends Component {
               render() {
                   return (
                    <View style={styles.container}>
                        <Image source={this.props.image} style={styles.image} />
                        <Autor email='fagnertakeshi@gmail.com' nickname='Takeshi'/> 
                        <Comments comments={this.props.comments} />
                    </View>
                   )
               }
           }

const styles = StyleSheet.create({
    container: {
        flex:1,
    },
    image:{
        height:Dimensions.get('window').width,
        width:Dimensions.get('window').width *3/4,
        resizeMode:'contain'
    },
})

export default Post
