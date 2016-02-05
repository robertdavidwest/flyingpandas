import pandas as pd
from copy import deepcopy

import time

def print_merge_stats(df, t_mergevar, mergetype, how, _left_on, _right_on,
                      matches_required, noprint, msg):
            m_stats = pd.DataFrame(pd.value_counts(df[t_mergevar]),
                                   columns=['Count'])
            if noprint == False:
                print('Merge Statistics' + ' (' + how + '  ' + mergetype + ')')
                print('On: ' + str(_left_on) + '/' + str(_right_on))
                print(m_stats)

            if (m_stats.loc['both', 'Count'][0] == 0) & matches_required:
                raise AssertionError('mer_9', 'Intersection of left Frame and '
                                              'right Frame is empty')

def merge(mergetype, left_frame, right_frame, how='invalid', sets=None,
          on=None, left_on=None, right_on=None, left_index=False,
          right_index=False, sort=False, suffixes=('__', '__'),
          copy=True, indicator=False, matches_required=True, noprint=False,
          msg=''):


    """
    This is a wrapper around pandas.merge with some extra functionality

    Parameters
    ----------
    mergetype : str
        '1:1', '1:m', 'm:1' or 'm:m'
        This variable describes the expectation of uniqueness in the key
        variable used in the merge.
        Examples:
        * if you expect both 'left_on' and right_on' to contain zero
        duplicates when you are merging two dataframes, then you would specify
        '1:1'.
        * if you were expecting 'left_on' to contain zero duplicates but
        'right_on' to contain some duplicates, you would specify '1:m'
        * if you were expecting 'right_on' to contain zero duplicates but
        'left_on' to contain some duplicates, you would specify 'm:1'
        * If you expected duplicates in both keys, you would specify 'm:m'.
         Using this option will give the standard result you would receive from
         using pandas.merge direct
    left : DataFrame
        same as pandas.merge
    right : DataFrame
        same as pandas.merge
    sets :  Optional[str or list of str]
        sets describe the expected set of values to be returned in the variable
        _merge in the merged dataframe.
        e.g.    *  'both'
                * ['left_only', 'both']
    matches_required: Optional [Bool]; default=True
        merge will exit if there are no observations which came from both
        the left and right DateFrame
    noprint: Optional [Bool]; Default=False
        suppress standard output from merge
    msg: Optional [str]
        print message at the top of merge result summary
        the --msg-- option works even with noprint==True


    For all other ``pandas.merge`` parameters see the ``pandas.merge`` docstring

    Returns
    -------
    merged : DataFrame
        if ``mergevar`` is specified then ``merged`` will include the variable
         ``mergevar``

    """


    if noprint==False:
        print
        print('-'*40)
    if msg:
        print(msg)
    __start = time.time()

    if how not in ['left', 'right', 'inner', 'outer']:
        raise AssertionError('mer_8', 'Invalid input for -how-')

    if not indicator:
        drop_t_mergevar = True
        t_mergevar = '_merge'
    else:
        drop_t_mergevar = False
        if type(indicator) == str:
            t_mergevar = indicator
        else:
            t_mergevar = '_merge'

    if ((on and left_on ) or (on and right_on)):
        raise AssertionError('mer_1', "cannot specify 'on' as well as "
                                      "'left_on' or as well as 'right_on'")
    if on:
        _left_on = on
        _right_on = on
    if left_on:
        _left_on = left_on

    if right_on:
        _right_on = right_on

    # if variables are str, convert to lists
    if type(_left_on) == str:
        _left_on = [_left_on]

    # if variables are str, convert to lists
    if type(_right_on) == str:
        _right_on = [_right_on]


    # if no suffixes specified, make sure there aren't columns with the same
    # names (excluding the key variables, of course)
    if suffixes == ('__', '__'):
        for ll in list(left_frame):
            for rr in list(right_frame):
                if ((ll == rr) & (ll not in _left_on) & (rr not in _right_on)):
                    raise AssertionError('mer_10', 'Column -' + ll +
                                         '- exists in both dataframes you are '
                                         'trying to merge. Specify suffixes if '
                                         'both columns are to be kept.')

    # check compatibility of key column data types
    # distinguish between  three types: object/datetime64[ns]/numeric
    for a, b in zip(_left_on, _right_on):
        lt = left_frame[a].dtypes
        rt = right_frame[b].dtypes
        if (lt != 'object') & (lt != 'datetime64[ns]'):
            lt = 'numeric'
        if (rt != 'object') & (rt != 'datetime64[ns]'):
            rt = 'numeric'
        if str(lt) != str(rt):
            err_msg = 'Column -' + a + '- in Left dataframe is ' + str(lt) + \
                    '; Column -' + b + '- in Right dataframe is ' + str(rt)
            raise AssertionError('mer_2', err_msg)

    if mergetype not in ['1:1', '1:m', 'm:1', 'm:m']:
        raise AssertionError('mer_3', 'mergetype needs to be '
                                      '1:1, 1:m, m:1, or m:m')

    # Check for uniqueness of data by key variable
    if mergetype[0] == '1':
        if sum(left_frame[_left_on].duplicated()) > 0:
            raise AssertionError('mer_4', 'Left key is not unique')
    if mergetype[-1] == '1':
        if sum(right_frame[_right_on].duplicated()) > 0:
            raise AssertionError('mer_5', 'Right key is not unique')



    if sets:
        new_frame = pd.merge(left_frame, right_frame, how='outer', on=on,
                             left_on=left_on, right_on=right_on,
                             left_index=left_index, right_index=right_index,
                             sort=sort, suffixes=suffixes, copy=copy,
                             indicator=t_mergevar)

        if type(sets) == 'str':
            sets = [sets]
        for set in sets:
            if set not in ['left_only', 'right_only', 'both']:
                raise AssertionError('mer_6', 'Sets must only contain '
                                              '"left_only", "right_only" or '
                                              '"both" as inputs')

        n_problems = (~new_frame[t_mergevar].isin(sets)).sum()
        if n_problems > 0:
            print('-'*100)
            print_merge_stats(new_frame, t_mergevar, mergetype, how, _left_on,
                              _right_on, matches_required, noprint, msg)
            print
            print('Examples of cases violating -sets- condition')
            print(new_frame.loc[~new_frame[t_mergevar].isin(sets)].head(10))
            raise AssertionError('mer_7', 'Not all observations from '
                                          'specified sets')
        else:
            # restrict rows based on --how--
            if how == 'left':
                set = ['left_only', 'both']
                new_frame = new_frame.loc[new_frame[t_mergevar].isin(set)]
            elif how == 'right':
                set = ['right_only', 'both']
                new_frame = new_frame.loc[new_frame[t_mergevar].isin(set)]
            elif how == 'inner':
                set = ['both']
                new_frame = new_frame.loc[new_frame[t_mergevar].isin(set)]
            elif how == 'outer':
                print
                # do nothing
            else:
                raise AssertionError('mer_8', 'Invalid input for -how-')

            print_merge_stats(new_frame, t_mergevar, mergetype, how, _left_on,
                              _right_on, matches_required, noprint, msg)

    else:
        # no check on the sets, do not necessarily do an outer merge
        new_frame = pd.merge(left_frame, right_frame, how=how, on=on,
                             left_on=left_on, right_on=right_on,
                             left_index=left_index, right_index=right_index,
                             sort=sort, suffixes=suffixes, copy=copy,
                             indicator=t_mergevar)

        print_merge_stats(new_frame, t_mergevar, mergetype, how, _left_on,
                          _right_on, matches_required, noprint, msg)

    #----------------------------------------------------
    # Some sanity checks based on number of observations
    l = len(left_frame)
    r = len(right_frame)
    n = len(new_frame)

    if how == 'left':
        if mergetype in ['m:1', '1:1']:
            assert n == l
        elif mergetype in ['1:m', 'm:m']:
            assert n >= l
            if mergetype == '1:m':
                assert n <= l + r - 1
            if mergetype == 'm:m':
                assert n <= l*r

    if how == 'right':
        if mergetype in ['1:m', '1:1']:
            assert n == r
        elif mergetype in ['m:1', 'm:m']:
            assert n >= r
            if mergetype == 'm:1':
                assert n <= l + r - 1
            if mergetype == 'm:m':
                assert n <= l*r

    if (how == 'inner'):
        if ('mergetype' == '1:1'):
            assert n <= max(l, r)

    if how == 'outer':
        assert n >= max(l, r)
        if mergetype in ['1:1', '1:m', 'm:1']:
            assert n <= l + r
        elif mergetype == 'm:m':
            assert n <= l*r

    if drop_t_mergevar:
        new_frame = new_frame.drop(axis=1, labels=t_mergevar)
    if noprint==False:
        print
        print 'Rows of   left dataframe:', l
        print 'Rows of  right dataframe:', r
        print 'Rows of merged dataframe:', n
        print
        print('merge time: (seconds)')
        print('{:5.3f}'.format(time.time() - __start))
        print('-'*40)

    return new_frame
