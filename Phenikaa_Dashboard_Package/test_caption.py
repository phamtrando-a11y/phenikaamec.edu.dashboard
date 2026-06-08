import pandas as pd
import dataframe_image as dfi

df = pd.DataFrame({'A': [1,2], 'B': [3,4]})
styler = df.style.set_caption("My Awesome Caption").set_table_styles([{
    'selector': 'caption',
    'props': [('font-size', '20px'), ('font-weight', 'bold')]
}])
dfi.export(styler, 'test.png', table_conversion='matplotlib')
